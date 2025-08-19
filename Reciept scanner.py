"""
Receipt Processing System for Australian Tax Categories
IMPROVED VERSION - Addresses code review feedback

Requirements:
- pip install pytesseract pillow openpyxl python-dateutil fuzzywuzzy python-levenshtein
- Install Tesseract OCR system: https://github.com/tesseract-ocr/tesseract
"""

import os
import re
import hashlib
import logging
import json
from datetime import datetime, date
from typing import List, Dict, Tuple, Optional, Set
from dataclasses import dataclass, asdict
from pathlib import Path
import configparser

from PIL import Image, ImageEnhance
import pytesseract
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, NamedStyle
from openpyxl.utils import get_column_letter
from dateutil import parser as date_parser
from fuzzywuzzy import fuzz


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('receipt_processor.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


@dataclass
class ReceiptData:
    """Data structure to hold extracted receipt information"""
    file_path: str
    filename: str = ""
    date: Optional[date] = None
    amount: Optional[float] = None
    merchant: Optional[str] = None
    category: Optional[str] = None
    confidence: float = 0.0
    raw_text: str = ""
    image_hash: str = ""
    processing_errors: List[str] = None
    manual_review_required: bool = False
    
    def __post_init__(self):
        if self.processing_errors is None:
            self.processing_errors = []
        if not self.filename:
            self.filename = Path(self.file_path).name


class Config:
    """Configuration management for the receipt processor"""
    
    def __init__(self, config_file: str = "config.ini"):
        self.config = configparser.ConfigParser()
        self.config_file = config_file
        self.load_config()
    
    def load_config(self):
        """Load configuration from file or create defaults"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
        else:
            self.create_default_config()
    
    def create_default_config(self):
        """Create default configuration file"""
        self.config['OCR'] = {
            'tesseract_config': '--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,/$-: ',
            'max_image_width': '2000',
            'max_image_height': '2000',
            'supported_formats': 'jpg,jpeg,png'
        }
        
        self.config['PROCESSING'] = {
            'duplicate_threshold': '0.85',
            'confidence_threshold': '0.6',
            'max_file_size_mb': '10'
        }
        
        self.config['OUTPUT'] = {
            'excel_filename': 'receipt_analysis.xlsx',
            'backup_enabled': 'true',
            'date_format': '%d/%m/%Y'
        }
        
        with open(self.config_file, 'w') as f:
            self.config.write(f)
        logger.info(f"Created default configuration file: {self.config_file}")
    
    def get(self, section: str, key: str, fallback=None):
        """Get configuration value with fallback"""
        return self.config.get(section, key, fallback=fallback)
    
    def getfloat(self, section: str, key: str, fallback=0.0):
        """Get float configuration value"""
        return self.config.getfloat(section, key, fallback=fallback)
    
    def getint(self, section: str, key: str, fallback=0):
        """Get integer configuration value"""
        return self.config.getint(section, key, fallback=fallback)
    
    def getboolean(self, section: str, key: str, fallback=False):
        """Get boolean configuration value"""
        return self.config.getboolean(section, key, fallback=fallback)


class ValidationError(Exception):
    """Custom exception for data validation errors"""
    pass


class DataValidator:
    """Validates and cleans extracted data"""
    
    @staticmethod
    def validate_amount(amount_str: str) -> Optional[float]:
        """Validate and parse amount string"""
        try:
            if not amount_str:
                return None
            
            # Clean amount string
            cleaned = re.sub(r'[^\d.]', '', amount_str)
            if not cleaned:
                return None
            
            amount = float(cleaned)
            
            # Reasonable range check (1 cent to $10,000)
            if 0.01 <= amount <= 10000.00:
                return round(amount, 2)
            else:
                logger.warning(f"Amount {amount} outside reasonable range")
                return None
                
        except (ValueError, TypeError):
            logger.warning(f"Could not parse amount: {amount_str}")
            return None
    
    @staticmethod
    def validate_date(date_str: str) -> Optional[date]:
        """Validate and parse date string"""
        try:
            if not date_str:
                return None
            
            # Try multiple parsing approaches
            parsed_date = date_parser.parse(date_str, fuzzy=True, dayfirst=True).date()
            
            # Reasonable range check (not future, not too old)
            today = date.today()
            if parsed_date <= today and (today - parsed_date).days <= 3650:  # 10 years
                return parsed_date
            else:
                logger.warning(f"Date {parsed_date} outside reasonable range")
                return None
                
        except (ValueError, TypeError):
            logger.warning(f"Could not parse date: {date_str}")
            return None
    
    @staticmethod
    def clean_merchant_name(merchant: str) -> Optional[str]:
        """Clean and normalize merchant name"""
        if not merchant:
            return None
        
        # Remove common suffixes and clean
        cleaned = re.sub(r'\b(PTY\s*LTD|LIMITED|LTD|INC|CORP)\b', '', merchant, flags=re.IGNORECASE)
        cleaned = re.sub(r'[^\w\s]', ' ', cleaned)
        cleaned = ' '.join(cleaned.split())  # Normalize whitespace
        
        return cleaned.strip()[:50] if cleaned.strip() else None


class OCRProcessor:
    """Handles image preprocessing and OCR text extraction with robust error handling"""
    
    def __init__(self, config: Config):
        self.config = config
        self.tesseract_config = config.get('OCR', 'tesseract_config')
        self.max_width = config.getint('OCR', 'max_image_width')
        self.max_height = config.getint('OCR', 'max_image_height')
        self.supported_formats = set(
            fmt.strip().lower() 
            for fmt in config.get('OCR', 'supported_formats').split(',')
        )
    
    def validate_image_file(self, image_path: str) -> bool:
        """Validate image file before processing"""
        try:
            path = Path(image_path)
            
            # Check file exists
            if not path.exists():
                logger.error(f"File does not exist: {image_path}")
                return False
            
            # Check file extension
            if path.suffix.lower().lstrip('.') not in self.supported_formats:
                logger.error(f"Unsupported format: {path.suffix}")
                return False
            
            # Check file size
            max_size_mb = self.config.getfloat('PROCESSING', 'max_file_size_mb')
            size_mb = path.stat().st_size / (1024 * 1024)
            if size_mb > max_size_mb:
                logger.error(f"File too large: {size_mb:.1f}MB > {max_size_mb}MB")
                return False
            
            return True
            
        except Exception as e:
            logger.error(f"Error validating image file {image_path}: {str(e)}")
            return False
    
    def preprocess_image(self, image_path: str) -> Optional[Image.Image]:
        """Preprocess image for better OCR results with error handling"""
        try:
            with Image.open(image_path) as img:
                # Convert to RGB if necessary
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Resize if too large
                if img.width > self.max_width or img.height > self.max_height:
                    img.thumbnail((self.max_width, self.max_height), Image.Resampling.LANCZOS)
                    logger.info(f"Resized image {image_path}")
                
                # Enhance contrast and sharpness for better OCR
                enhancer = ImageEnhance.Contrast(img)
                img = enhancer.enhance(1.2)
                
                enhancer = ImageEnhance.Sharpness(img)
                img = enhancer.enhance(1.1)
                
                return img.copy()  # Return copy to ensure original is closed
                
        except Exception as e:
            logger.error(f"Error preprocessing image {image_path}: {str(e)}")
            return None
    
    def extract_text(self, image_path: str) -> Tuple[str, List[str]]:
        """Extract text from receipt image with comprehensive error handling"""
        errors = []
        
        try:
            if not self.validate_image_file(image_path):
                return "", ["Invalid image file"]
            
            processed_image = self.preprocess_image(image_path)
            if processed_image is None:
                return "", ["Failed to preprocess image"]
            
            # Extract text using Tesseract
            text = pytesseract.image_to_string(processed_image, config=self.tesseract_config)
            
            if not text.strip():
                errors.append("No text extracted from image")
                logger.warning(f"No text extracted from {image_path}")
            
            return text.strip(), errors
            
        except pytesseract.TesseractNotFoundError:
            error_msg = "Tesseract OCR not found. Please install Tesseract OCR."
            errors.append(error_msg)
            logger.error(error_msg)
            return "", errors
            
        except Exception as e:
            error_msg = f"OCR extraction failed: {str(e)}"
            errors.append(error_msg)
            logger.error(f"Error extracting text from {image_path}: {str(e)}")
            return "", errors
    
    def calculate_image_hash(self, image_path: str) -> str:
        """Calculate perceptual hash of image for duplicate detection"""
        try:
            with Image.open(image_path) as img:
                # Convert to grayscale and resize for consistent hashing
                img = img.convert('L').resize((8, 8))
                pixels = list(img.getdata())
                
                # Calculate average
                avg = sum(pixels) / len(pixels)
                
                # Create hash based on pixels above/below average
                hash_bits = ''.join('1' if pixel > avg else '0' for pixel in pixels)
                return hash_bits
                
        except Exception as e:
            logger.error(f"Error calculating image hash for {image_path}: {str(e)}")
            return ""


class TextParser:
    """Enhanced text parser with better pattern matching and validation"""
    
    def __init__(self):
        # Australian date patterns (more comprehensive)
        self.date_patterns = [
            r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b',  # DD/MM/YYYY
            r'\b(\d{2,4}[/-]\d{1,2}[/-]\d{1,2})\b',  # YYYY/MM/DD
            r'\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4})\b',
            r'\b(\d{1,2}\.?\d{1,2}\.?\d{2,4})\b'     # DD.MM.YYYY
        ]
        
        # Enhanced currency patterns
        self.amount_patterns = [
            r'TOTAL[:\s]*\$?(\d+\.?\d{0,2})',        # TOTAL: $XX.XX (highest priority)
            r'(?:AMOUNT|AMT)[:\s]*\$?(\d+\.?\d{0,2})', # AMOUNT: $XX.XX
            r'\$\s*(\d+\.?\d{0,2})\s*(?:AUD|AU)?',    # $XX.XX AUD
            r'(\d+\.\d{2})\s*(?:AUD|AU\$|\$)',       # XX.XX AUD
            r'\$(\d+\.?\d{0,2})(?:\s|$)',            # $XX.XX at end of line
        ]
    
    def parse_date(self, text: str) -> Optional[date]:
        """Extract and validate date from OCR text"""
        for pattern in self.date_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                validated_date = DataValidator.validate_date(match)
                if validated_date:
                    return validated_date
        return None
    
    def parse_amount(self, text: str) -> Optional[float]:
        """Extract and validate amount with priority ordering"""
        amounts = []
        
        for pattern in self.amount_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                validated_amount = DataValidator.validate_amount(match)
                if validated_amount:
                    amounts.append(validated_amount)
        
        # Return the largest reasonable amount found
        return max(amounts) if amounts else None
    
    def parse_merchant(self, text: str) -> Optional[str]:
        """Extract merchant name from first few lines of receipt"""
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        if not lines:
            return None
        
        # Look at first 3 lines for merchant name
        for line in lines[:3]:
            # Skip lines that look like addresses or phone numbers
            if re.search(r'\d{4,}|\d+\s+\w+\s+(?:st|street|rd|road|ave|avenue)', line, re.IGNORECASE):
                continue
            
            # Clean and validate merchant name
            merchant = DataValidator.clean_merchant_name(line)
            if merchant and len(merchant) > 3:  # Must be meaningful length
                return merchant
        
        return None


class ATOClassifier:
    """Enhanced classifier with fuzzy matching and better categorization"""
    
    def __init__(self):
        # ATO expense categories with expanded keywords
        self.categories = {
            'Motor vehicle expenses': {
                'keywords': [
                    'bp', 'shell', 'caltex', '7-eleven', 'united', 'mobil', 'ampol',
                    'petrol', 'fuel', 'servo', 'service station', 'mechanics', 'car wash',
                    'parking', 'toll', 'rego', 'registration', 'roadworthy', 'tyres',
                    'automotive', 'auto', 'vehicle', 'car service'
                ],
                'merchants': [
                    'bp australia', 'shell australia', 'caltex', 'ampol', '7-eleven',
                    'united petroleum', 'liberty oil', 'puma energy'
                ]
            },
            'Travel expenses': {
                'keywords': [
                    'qantas', 'virgin', 'jetstar', 'tigerair', 'hotel', 'motel',
                    'accommodation', 'taxi', 'uber', 'airport', 'flight', 'airfare',
                    'booking.com', 'expedia', 'travel', 'train', 'bus'
                ],
                'merchants': [
                    'qantas airways', 'virgin australia', 'jetstar airways',
                    'uber', 'taxi', 'hotel', 'motel'
                ]
            },
            'Meals and entertainment (50% deductible)': {
                'keywords': [
                    'restaurant', 'cafe', 'mcdonald', 'kfc', 'subway', 'dominos',
                    'pizza', 'pub', 'bar', 'catering', 'lunch', 'dinner', 'food',
                    'hungry jacks', 'red rooster', 'nandos'
                ],
                'merchants': [
                    'mcdonalds', 'kfc', 'subway', 'dominos pizza', 'pizza hut',
                    'hungry jacks', 'red rooster', 'nandos'
                ]
            },
            'Office supplies and software': {
                'keywords': [
                    'officeworks', 'staples', 'harvey norman', 'jb hi-fi',
                    'software', 'microsoft', 'adobe', 'stationery', 'printer',
                    'ink', 'paper', 'computer', 'laptop', 'office', 'supplies'
                ],
                'merchants': [
                    'officeworks', 'staples', 'harvey norman', 'jb hi fi',
                    'microsoft', 'adobe systems'
                ]
            },
            'Professional development': {
                'keywords': [
                    'training', 'course', 'seminar', 'conference', 'workshop',
                    'education', 'university', 'tafe', 'certification', 'learning'
                ],
                'merchants': []
            },
            'Insurance': {
                'keywords': [
                    'insurance', 'aami', 'allianz', 'suncorp', 'nrma', 'rac',
                    'professional indemnity', 'public liability', 'cover'
                ],
                'merchants': [
                    'aami', 'allianz', 'suncorp', 'nrma insurance', 'rac'
                ]
            },
            'Repairs and maintenance': {
                'keywords': [
                    'repair', 'maintenance', 'service', 'fix', 'replacement',
                    'plumber', 'electrician', 'handyman'
                ],
                'merchants': []
            },
            'Advertising and marketing': {
                'keywords': [
                    'advertising', 'marketing', 'facebook', 'google ads', 'meta',
                    'promotion', 'flyers', 'business cards', 'website', 'seo'
                ],
                'merchants': [
                    'facebook', 'google', 'meta'
                ]
            },
            'Bank fees': {
                'keywords': [
                    'bank fee', 'transaction fee', 'account fee', 'overdraft',
                    'atm fee', 'monthly fee'
                ],
                'merchants': [
                    'commonwealth bank', 'anz', 'westpac', 'nab'
                ]
            },
            'Other deductible expenses': {
                'keywords': [],
                'merchants': []
            }
        }
    
    def classify_expense(self, merchant: str, raw_text: str) -> Tuple[str, float]:
        """Classify expense using fuzzy matching and multiple signals"""
        if not merchant and not raw_text:
            return 'Other deductible expenses', 0.1
        
        best_category = 'Other deductible expenses'
        best_score = 0.0
        
        search_text = f"{merchant or ''} {raw_text}".lower()
        
        for category, data in self.categories.items():
            if category == 'Other deductible expenses':
                continue
                
            category_score = 0.0
            total_checks = 0
            
            # Check merchant fuzzy matching
            if merchant:
                for known_merchant in data['merchants']:
                    similarity = fuzz.partial_ratio(merchant.lower(), known_merchant.lower())
                    if similarity > 70:  # 70% similarity threshold
                        category_score += similarity / 100
                        total_checks += 1
            
            # Check keyword matching
            for keyword in data['keywords']:
                if keyword in search_text:
                    category_score += 1.0
                    total_checks += 1
                else:
                    # Fuzzy keyword matching
                    for word in search_text.split():
                        if fuzz.ratio(keyword, word) > 80:  # 80% similarity for keywords
                            category_score += 0.5
                            total_checks += 1
                            break
            
            # Calculate final score
            if total_checks > 0:
                final_score = min(category_score / max(total_checks, 1), 1.0)
                if final_score > best_score:
                    best_score = final_score
                    best_category = category
        
        # Ensure minimum confidence
        confidence = max(best_score, 0.1 if best_category != 'Other deductible expenses' else 0.05)
        
        return best_category, confidence


class DuplicateDetector:
    """Enhanced duplicate detection with multiple algorithms"""
    
    def __init__(self, config: Config):
        self.config = config
        self.processed_receipts: List[ReceiptData] = []
        self.similarity_threshold = config.getfloat('PROCESSING', 'duplicate_threshold')
    
    def is_duplicate(self, receipt: ReceiptData) -> Tuple[bool, Optional[ReceiptData], float]:
        """Check for duplicates using multiple methods"""
        best_match = None
        highest_similarity = 0.0
        
        for existing in self.processed_receipts:
            similarity = self.calculate_similarity(receipt, existing)
            
            if similarity > highest_similarity:
                highest_similarity = similarity
                best_match = existing
        
        is_duplicate = highest_similarity >= self.similarity_threshold
        return is_duplicate, best_match, highest_similarity
    
    def calculate_similarity(self, receipt1: ReceiptData, receipt2: ReceiptData) -> float:
        """Calculate comprehensive similarity score"""
        scores = []
        weights = []
        
        # Image hash comparison (highest weight)
        if receipt1.image_hash and receipt2.image_hash:
            hash_similarity = self.hamming_similarity(receipt1.image_hash, receipt2.image_hash)
            scores.append(hash_similarity)
            weights.append(0.4)
        
        # Date comparison
        if receipt1.date and receipt2.date:
            date_diff = abs((receipt1.date - receipt2.date).days)
            date_similarity = max(0, 1 - (date_diff / 7))  # Same week = high similarity
            scores.append(date_similarity)
            weights.append(0.25)
        
        # Amount comparison
        if receipt1.amount and receipt2.amount:
            amount_diff = abs(receipt1.amount - receipt2.amount)
            amount_similarity = max(0, 1 - (amount_diff / max(receipt1.amount, receipt2.amount)))
            scores.append(amount_similarity)
            weights.append(0.25)
        
        # Merchant comparison
        if receipt1.merchant and receipt2.merchant:
            merchant_similarity = fuzz.ratio(receipt1.merchant, receipt2.merchant) / 100
            scores.append(merchant_similarity)
            weights.append(0.1)
        
        # Calculate weighted average
        if scores:
            total_weight = sum(weights)
            weighted_sum = sum(score * weight for score, weight in zip(scores, weights))
            return weighted_sum / total_weight
        
        return 0.0
    
    def hamming_similarity(self, hash1: str, hash2: str) -> float:
        """Calculate similarity between two binary hashes"""
        if len(hash1) != len(hash2) or not hash1 or not hash2:
            return 0.0
        
        differences = sum(c1 != c2 for c1, c2 in zip(hash1, hash2))
        return 1 - (differences / len(hash1))
    
    def add_processed_receipt(self, receipt: ReceiptData):
        """Add receipt to processed list for future duplicate checking"""
        self.processed_receipts.append(receipt)


class ExcelManager:
    """Enhanced Excel management with better formatting and error handling"""
    
    def __init__(self, config: Config):
        self.config = config
        self.output_file = config.get('OUTPUT', 'excel_filename')
        self.backup_enabled = config.getboolean('OUTPUT', 'backup_enabled')
        self.workbook = None
        self.setup_workbook()
    
    def setup_workbook(self):
        """Create Excel workbook with professional formatting"""
        try:
            # Backup existing file if it exists
            if self.backup_enabled and os.path.exists(self.output_file):
                backup_name = f"{self.output_file}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                os.rename(self.output_file, backup_name)
                logger.info(f"Backed up existing file to: {backup_name}")
            
            self.workbook = openpyxl.Workbook()
            self.workbook.remove(self.workbook.active)  # Remove default sheet
            
            # Create custom styles
            self.create_styles()
            
            # Create all sheets
            self.create_main_sheet()
            self.create_review_sheet()
            self.create_duplicates_sheet()
            self.create_summary_sheet()
            
            logger.info("Excel workbook initialized successfully")
            
        except Exception as e:
            logger.error(f"Error setting up Excel workbook: {str(e)}")
            raise
    
    def create_styles(self):
        """Create named styles for consistent formatting"""
        # Header style
        header_style = NamedStyle(name="header")
        header_style.font = Font(bold=True, color="FFFFFF")
        header_style.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_style.alignment = Alignment(horizontal="center", vertical="center")
        
        # Currency style
        currency_style = NamedStyle(name="currency", number_format="$#,##0.00")
        
        # Date style
        date_style = NamedStyle(name="date", number_format="DD/MM/YYYY")
        
        # Add styles to workbook
        try:
            self.workbook.add_named_style(header_style)
            self.workbook.add_named_style(currency_style)
            self.workbook.add_named_style(date_style)
        except ValueError:
            # Styles already exist
            pass
    
    def create_main_sheet(self):
        """Create main processed receipts sheet"""
        ws = self.workbook.create_sheet("Processed Receipts")
        headers = ["Date", "Amount", "Merchant", "Category", "Confidence", "Filename", "Processing Notes"]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.style = "header"
        
        # Auto-adjust column widths
        column_widths = [12, 15, 25, 30, 12, 20, 30]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        
        # Freeze header row
        ws.freeze_panes = "A2"
        
        logger.info("Created main processed receipts sheet")
    
    def create_review_sheet(self):
        """Create review sheet for manual verification"""
        ws = self.workbook.create_sheet("Needs Review")
        headers = ["Date", "Amount", "Merchant", "Suggested Category", "Confidence", 
                  "Manual Category", "Filename", "Issues"]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.style = "header"
        
        # Add data validation for manual category column
        from openpyxl.worksheet.datavalidation import DataValidation
        
        categories = list(ATOClassifier().categories.keys())
        dv = DataValidation(type="list", formula1=f'"{",".join(categories)}"', allow_blank=True)
        dv.prompt = "Select a category"
        dv.promptTitle = "Category Selection"
        ws.add_data_validation(dv)
        
        # Apply validation to manual category column (F)
        dv.add(f"F2:F1000")
        
        logger.info("Created review sheet with data validation")
    
    def create_duplicates_sheet(self):
        """Create duplicates sheet"""
        ws = self.workbook.create_sheet("Potential Duplicates")
        headers = ["Original File", "Duplicate File", "Similarity Score", "Original Date", 
                  "Duplicate Date", "Original Amount", "Duplicate Amount", "Action Required"]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.style = "header"
        
        logger.info("Created duplicates sheet")
    
    def create_summary_sheet(self):
        """Create summary sheet with category totals"""
        ws = self.workbook.create_sheet("Summary by Category")
        headers = ["Category", "Total Amount", "Number of Receipts", "Average Amount", 
                  "Tax Deductible Amount", "Notes"]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.style = "header"
        
        # Add note about meals entertainment 50% rule
        ws.cell(row=2, column=6, value="Meals & Entertainment: Only 50% deductible for tax")
        
        logger.info("Created summary sheet")
    
    def add_receipt(self, receipt: ReceiptData, sheet_name: str = "Processed Receipts"):
        """Add receipt data to specified sheet with proper formatting"""
        try:
            ws = self.workbook[sheet_name]
            row = ws.max_row + 1
            
            # Prepare data
            date_str = receipt.date.strftime(self.config.get('OUTPUT', 'date_format')) if receipt.date else ""
            errors_str = "; ".join(receipt.processing_errors) if receipt.processing_errors else ""
            
            data = [
                date_str,
                receipt.amount or 0,
                receipt.merchant or "",
                receipt.category or "",
                receipt.confidence,
                receipt.filename,
                errors_str
            ]
            
            # Write data with formatting
            for col, value in enumerate(data, 1):
                cell = ws.cell(row=row, column=col, value=value)
                
                # Apply specific formatting
                if col == 1 and receipt.date:  # Date column
                    cell.style = "date"
                elif col == 2:  # Amount column
                    cell.style = "currency"
                elif col == 5:  # Confidence column
                    cell.number_format = "0.00"
            
            logger.debug(f"Added receipt {receipt.filename} to {sheet_name}")
            
        except Exception as e:
            logger.error(f"Error adding receipt to Excel: {str(e)}")
    
    def add_duplicate(self, original: ReceiptData, duplicate: ReceiptData, similarity: float):
        """Add duplicate pair to duplicates sheet"""
        try:
            ws = self.workbook["Potential Duplicates"]
            row = ws.max_row + 1
            
            data = [
                original.filename,
                duplicate.filename,
                similarity,
                original.date.strftime(self.config.get('OUTPUT', 'date_format')) if original.date else "",
                duplicate.date.strftime(self.config.get('OUTPUT', 'date_format')) if duplicate.date else "",
                original.amount or 0,
                duplicate.amount or 0,
                "Review and delete one" if similarity > 0.9 else "Verify manually"
            ]
            
            for col, value in enumerate(data, 1):
                cell = ws.cell(row=row, column=col, value=value)
                if col in [3]:  # Similarity score
                    cell.number_format = "0.00"
                elif col in [4, 5]:  # Date columns
                    cell.style = "date"
                elif col in [6, 7]:  # Amount columns
                    cell.style = "currency"
            
            logger.debug(f"Added duplicate pair: {original.filename} / {duplicate.filename}")
            
        except Exception as e:
            logger.error(f"Error adding duplicate to Excel: {str(e)}")
    
    def update_summary(self, processed_receipts: List[ReceiptData]):
        """Update summary sheet with category totals"""
        try:
            ws = self.workbook["Summary by Category"]
            
            # Clear existing data (keep headers)
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.value = None
            
            # Calculate totals by category
            category_totals = {}
            for receipt in processed_receipts:
                if receipt.category and receipt.amount:
                    if receipt.category not in category_totals:
                        category_totals[receipt.category] = {'total': 0, 'count': 0}
                    category_totals[receipt.category]['total'] += receipt.amount
                    category_totals[receipt.category]['count'] += 1
            
            # Write summary data
            row = 2
            for category, data in sorted(category_totals.items()):
                total_amount = data['total']
                count = data['count']
                average = total_amount / count if count > 0 else 0
                
                # Calculate tax deductible amount
                if "50% deductible" in category.lower():
                    deductible = total_amount * 0.5
                else:
                    deductible = total_amount
                
                ws.cell(row=row, column=1, value=category)
                ws.cell(row=row, column=2, value=total_amount).style = "currency"
                ws.cell(row=row, column=3, value=count)
                ws.cell(row=row, column=4, value=average).style = "currency"
                ws.cell(row=row, column=5, value=deductible).style = "currency"
                
                row += 1
            
            logger.info("Updated summary sheet with category totals")
            
        except Exception as e:
            logger.error(f"Error updating summary sheet: {str(e)}")
    
    def save_workbook(self):
        """Save the Excel workbook with error handling"""
        try:
            self.workbook.save(self.output_file)
            logger.info(f"Excel workbook saved successfully: {self.output_file}")
        except PermissionError:
            error_msg = f"Cannot save {self.output_file} - file may be open in Excel"
            logger.error(error_msg)
            raise PermissionError(error_msg)
        except Exception as e:
            logger.error(f"Error saving Excel workbook: {str(e)}")
            raise


class BatchProcessor:
    """Enhanced main processor with comprehensive error handling and progress tracking"""
    
    def __init__(self, input_folder: str, config_file: str = "config.ini"):
        self.input_folder = input_folder
        self.config = Config(config_file)
        
        # Initialize components
        self.ocr_processor = OCRProcessor(self.config)
        self.text_parser = TextParser()
        self.classifier = ATOClassifier()
        self.duplicate_detector = DuplicateDetector(self.config)
        self.data_validator = DataValidator()
        
        # Initialize Excel manager
        try:
            self.excel_manager = ExcelManager(self.config)
        except Exception as e:
            logger.error(f"Failed to initialize Excel manager: {str(e)}")
            raise
        
        # Processing statistics
        self.stats = {
            'total_files': 0,
            'processed': 0,
            'duplicates': 0,
            'errors': 0,
            'needs_review': 0,
            'skipped': 0
        }
        
        # Store processed receipts
        self.processed_receipts: List[ReceiptData] = []
        self.review_receipts: List[ReceiptData] = []
        self.duplicate_pairs: List[Tuple[ReceiptData, ReceiptData, float]] = []
    
    def get_image_files(self) -> List[str]:
        """Get all supported image files from input folder"""
        try:
            supported_extensions = {'.jpg', '.jpeg', '.png'}
            image_files = []
            
            input_path = Path(self.input_folder)
            if not input_path.exists():
                raise FileNotFoundError(f"Input folder does not exist: {self.input_folder}")
            
            for file_path in input_path.rglob("*"):
                if file_path.is_file() and file_path.suffix.lower() in supported_extensions:
                    image_files.append(str(file_path))
            
            logger.info(f"Found {len(image_files)} image files to process")
            return sorted(image_files)  # Sort for consistent processing order
            
        except Exception as e:
            logger.error(f"Error scanning input folder: {str(e)}")
            return []
    
    def process_folder(self):
        """Main method to process entire folder with progress tracking"""
        logger.info("Starting batch processing of receipt folder")
        
        try:
            image_files = self.get_image_files()
            if not image_files:
                logger.warning("No image files found to process")
                return
            
            self.stats['total_files'] = len(image_files)
            
            for i, file_path in enumerate(image_files, 1):
                logger.info(f"Processing file {i}/{len(image_files)}: {Path(file_path).name}")
                
                try:
                    receipt = self.process_single_receipt(file_path)
                    
                    if receipt is None:
                        self.stats['skipped'] += 1
                        continue
                    
                    # Check for duplicates
                    is_duplicate, original, similarity = self.duplicate_detector.is_duplicate(receipt)
                    
                    if is_duplicate and original:
                        logger.info(f"Duplicate detected: {receipt.filename} (similarity: {similarity:.2f})")
                        self.duplicate_pairs.append((original, receipt, similarity))
                        self.excel_manager.add_duplicate(original, receipt, similarity)
                        self.stats['duplicates'] += 1
                    else:
                        # Process as new receipt
                        self.duplicate_detector.add_processed_receipt(receipt)
                        
                        if self.needs_review(receipt):
                            self.review_receipts.append(receipt)
                            self.excel_manager.add_receipt(receipt, "Needs Review")
                            self.stats['needs_review'] += 1
                        else:
                            self.processed_receipts.append(receipt)
                            self.excel_manager.add_receipt(receipt, "Processed Receipts")
                            self.stats['processed'] += 1
                
                except Exception as e:
                    logger.error(f"Error processing {file_path}: {str(e)}")
                    self.stats['errors'] += 1
                    continue
            
            # Update summary sheet
            all_receipts = self.processed_receipts + self.review_receipts
            self.excel_manager.update_summary(all_receipts)
            
            # Save Excel file
            self.excel_manager.save_workbook()
            
            logger.info("Batch processing completed successfully")
            
        except Exception as e:
            logger.error(f"Error during batch processing: {str(e)}")
            raise
    
    def process_single_receipt(self, file_path: str) -> Optional[ReceiptData]:
        """Process a single receipt with comprehensive error handling"""
        try:
            # Initialize receipt data
            receipt = ReceiptData(file_path=file_path)
            
            # Extract text using OCR
            raw_text, ocr_errors = self.ocr_processor.extract_text(file_path)
            receipt.raw_text = raw_text
            receipt.processing_errors.extend(ocr_errors)
            
            if not raw_text:
                receipt.processing_errors.append("No text could be extracted")
                return receipt
            
            # Calculate image hash for duplicate detection
            receipt.image_hash = self.ocr_processor.calculate_image_hash(file_path)
            
            # Parse structured data
            receipt.date = self.text_parser.parse_date(raw_text)
            receipt.amount = self.text_parser.parse_amount(raw_text)
            receipt.merchant = self.text_parser.parse_merchant(raw_text)
            
            # Validate critical data
            if not receipt.date:
                receipt.processing_errors.append("Could not extract date")
            if not receipt.amount:
                receipt.processing_errors.append("Could not extract amount")
            if not receipt.merchant:
                receipt.processing_errors.append("Could not identify merchant")
            
            # Classify expense
            category, confidence = self.classifier.classify_expense(receipt.merchant, raw_text)
            receipt.category = category
            receipt.confidence = confidence
            
            # Mark for review if low confidence or missing critical data
            if confidence < self.config.getfloat('PROCESSING', 'confidence_threshold'):
                receipt.processing_errors.append(f"Low classification confidence: {confidence:.2f}")
            
            logger.debug(f"Processed receipt: {receipt.filename} -> {category} ({confidence:.2f})")
            return receipt
            
        except Exception as e:
            logger.error(f"Error processing receipt {file_path}: {str(e)}")
            error_receipt = ReceiptData(file_path=file_path)
            error_receipt.processing_errors.append(f"Processing failed: {str(e)}")
            return error_receipt
    
    def needs_review(self, receipt: ReceiptData) -> bool:
        """Determine if receipt needs manual review"""
        confidence_threshold = self.config.getfloat('PROCESSING', 'confidence_threshold')
        
        return (
            receipt.confidence < confidence_threshold or
            not receipt.date or
            not receipt.amount or
            len(receipt.processing_errors) > 0
        )
    
    def print_statistics(self):
        """Print comprehensive processing statistics"""
        print("\n" + "="*60)
        print("📊 PROCESSING STATISTICS")
        print("="*60)
        print(f"Total files found:     {self.stats['total_files']}")
        print(f"Successfully processed: {self.stats['processed']}")
        print(f"Needs manual review:   {self.stats['needs_review']}")
        print(f"Duplicates detected:   {self.stats['duplicates']}")
        print(f"Files with errors:     {self.stats['errors']}")
        print(f"Files skipped:         {self.stats['skipped']}")
        print("="*60)
        
        if self.processed_receipts:
            total_amount = sum(r.amount for r in self.processed_receipts if r.amount)
            print(f"Total processed amount: ${total_amount:,.2f}")
        
        if self.stats['needs_review'] > 0:
            print(f"\n⚠️  {self.stats['needs_review']} receipts need manual review")
        
        if self.stats['duplicates'] > 0:
            print(f"🔍 {self.stats['duplicates']} potential duplicates found")
    
    def generate_report(self):
        """Generate detailed processing report"""
        report_file = f"processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
        try:
            with open(report_file, 'w') as f:
                f.write("RECEIPT PROCESSING REPORT\n")
                f.write("=" * 50 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Input folder: {self.input_folder}\n")
                f.write(f"Output file: {self.excel_manager.output_file}\n\n")
                
                # Statistics
                f.write("PROCESSING STATISTICS:\n")
                for key, value in self.stats.items():
                    f.write(f"{key.replace('_', ' ').title()}: {value}\n")
                
                # Error summary
                if self.review_receipts:
                    f.write("\nFILES REQUIRING REVIEW:\n")
                    for receipt in self.review_receipts:
                        f.write(f"- {receipt.filename}: {'; '.join(receipt.processing_errors)}\n")
                
                # Duplicate summary
                if self.duplicate_pairs:
                    f.write("\nPOTENTIAL DUPLICATES:\n")
                    for original, duplicate, similarity in self.duplicate_pairs:
                        f.write(f"- {original.filename} / {duplicate.filename} (similarity: {similarity:.2f})\n")
            
            logger.info(f"Processing report saved: {report_file}")
            
        except Exception as e:
            logger.error(f"Error generating report: {str(e)}")


def check_dependencies():
    """Check if all required dependencies are available"""
    missing_deps = []
    
    try:
        import pytesseract
        # Test if Tesseract is actually installed
        pytesseract.get_tesseract_version()
    except Exception:
        missing_deps.append("Tesseract OCR (system installation required)")
    
    required_packages = [
        'PIL', 'openpyxl', 'dateutil', 'fuzzywuzzy'
    ]
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_deps.append(f"Python package: {package}")
    
    if missing_deps:
        print("❌ Missing dependencies:")
        for dep in missing_deps:
            print(f"   - {dep}")
        print("\nPlease install missing dependencies before running.")
        return False
    
    return True


def main():
    """Enhanced main function with better user interaction and error handling"""
    print("🧾 ADVANCED RECEIPT PROCESSING SYSTEM")
    print("📋 Australian Tax Office Categories")
    print("=" * 60)
    
    # Check dependencies first
    if not check_dependencies():
        return
    
    # Get input folder
    while True:
        input_folder = input("\n📁 Enter the folder path containing your receipts: ").strip()
        
        if input_folder.lower() in ['quit', 'exit', 'q']:
            print("Goodbye! 👋")
            return
        
        if os.path.exists(input_folder):
            break
        else:
            print(f"❌ Error: Folder '{input_folder}' does not exist!")
            print("💡 Tip: You can type 'quit' to exit")
    
    try:
        # Initialize processor
        print(f"\n🔄 Initializing processor...")
        processor = BatchProcessor(input_folder)
        
        # Show configuration
        print(f"📊 Configuration loaded:")
        print(f"   - Duplicate threshold: {processor.config.getfloat('PROCESSING', 'duplicate_threshold'):.2f}")
        print(f"   - Confidence threshold: {processor.config.getfloat('PROCESSING', 'confidence_threshold'):.2f}")
        print(f"   - Output file: {processor.excel_manager.output_file}")
        
        # Start processing
        print(f"\n🚀 Starting batch processing...")
        processor.process_folder()
        
        # Show results
        processor.print_statistics()
        processor.generate_report()
        
        print(f"\n✅ Processing complete!")
        print(f"📁 Results saved to: {processor.excel_manager.output_file}")
        print(f"📄 Processing report generated")
        
        print(f"\n📝 NEXT STEPS:")
        print(f"1. Open {processor.excel_manager.output_file} in Excel")
        print(f"2. Review the 'Needs Review' sheet for items requiring attention")
        print(f"3. Check the 'Potential Duplicates' sheet")
        print(f"4. Use the 'Summary by Category' sheet for your tax preparation")
        print(f"5. Save a copy of the file for your records")
        
        # Ask if user wants to open the file
        try:
            response = input(f"\n🔍 Would you like to open the Excel file now? (y/n): ").strip().lower()
            if response in ['y', 'yes']:
                import subprocess
                import platform
                
                if platform.system() == 'Windows':
                    os.startfile(processor.excel_manager.output_file)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', processor.excel_manager.output_file])
                else:  # Linux
                    subprocess.run(['xdg-open', processor.excel_manager.output_file])
        except:
            pass  # If opening fails, no big deal
        
    except KeyboardInterrupt:
        print(f"\n\n⏹️  Processing interrupted by user")
    except Exception as e:
        logger.error(f"Fatal error: {str(e)}")
        print(f"\n❌ A fatal error occurred: {str(e)}")
        print(f"📄 Check the log file 'receipt_processor.log' for more details")


if __name__ == "__main__":
    main()
