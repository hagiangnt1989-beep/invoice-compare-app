"""
Invoice PDF Parser

Handles both text-based and scanned (image-based) PDFs.
- Text-based PDFs: Direct text extraction
- Scanned PDFs: Convert to images and apply OCR
"""

from typing import Dict, Any, List
import fitz  # PyMuPDF alternative if available, otherwise use pdf2image
from pdf2image import convert_from_bytes
from PIL import Image
import io
import re
from parsers.invoice_ocr import (
    parse_invoice_image,
    extract_line_items_from_text,
    extract_totals_from_text,
    preprocess_image,
    get_ocr_engine
)
import numpy as np


def parse_invoice_pdf(file) -> Dict[str, Any]:
    """
    Parse PDF invoice.
    
    Strategy:
    1. Try text extraction first
    2. If text is minimal/poor, convert to images and use OCR
    
    Args:
        file: File-like object (uploaded PDF)
        
    Returns:
        Dictionary with 'line_items' list and 'totals' dict
    """
    try:
        # Read PDF bytes
        pdf_bytes = file.read()
        
        # First, try text extraction
        text_content = extract_text_from_pdf(pdf_bytes)
        
        # Check if text extraction was successful
        # (some scanned PDFs have minimal/garbage text)
        if is_text_extractable(text_content):
            # Parse structured text
            line_items = extract_line_items_from_text_content(text_content)
            totals = extract_totals_from_text_content(text_content)
            
            return {
                'line_items': line_items,
                'totals': totals
            }
        else:
            # Fallback to OCR (scanned PDF)
            return parse_scanned_pdf(pdf_bytes)
        
    except Exception as e:
        raise Exception(f"Error parsing PDF invoice: {str(e)}")


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """
    Extract text from PDF using PyMuPDF (fitz).
    Falls back to pdfplumber if needed.
    """
    try:
        # Try PyMuPDF first
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text
    except ImportError:
        # Fallback to pdfplumber
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""
                return text
        except ImportError:
            # If neither library available, return empty
            return ""


def is_text_extractable(text: str) -> bool:
    """
    Determine if extracted text is meaningful.
    Scanned PDFs often have minimal or garbage text.
    """
    # Remove whitespace
    clean_text = text.strip()
    
    # Check length
    if len(clean_text) < 50:
        return False
    
    # Check for actual words (not just garbage characters)
    words = re.findall(r'\b[a-zA-Zàáảãạăắằẳẵặâấầẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôốồổỗộơớờởỡợùúủũụưứừửữựỳýỷỹỵđ]{3,}\b', clean_text)
    
    if len(words) < 5:
        return False
    
    return True


def extract_line_items_from_text_content(text: str) -> List[Dict]:
    """
    Extract line items from raw PDF text.
    """
    line_items = []
    
    # Split into lines
    lines = text.split('\n')
    
    # Find table data (lines with product info and numbers)
    in_table = False
    header_found = False
    
    for line in lines:
        line_lower = line.lower()
        
        # Detect table header
        if not header_found and any(kw in line_lower for kw in [
            'sản phẩm', 'product', 'item', 'số lượng', 'quantity', 'thành tiền', 'amount'
        ]):
            in_table = True
            header_found = True
            continue
        
        # Skip until header found
        if not in_table:
            continue
        
        # Stop at summary section
        if any(kw in line_lower for kw in ['tổng cộng', 'tổng thanh toán', 'total', 'vat', 'thuế']):
            break
        
        # Extract item from line
        item = parse_line_item(line)
        if item:
            line_items.append(item)
    
    return line_items


def parse_line_item(line: str) -> Dict:
    """
    Parse a single line as a line item.
    Expected format: "Product name 50000 10 500000"
    (product, denomination, quantity, amount)
    """
    # Extract all numbers from line
    numbers = re.findall(r'[\d.,]+', line)
    
    if len(numbers) < 2:
        return None
    
    # Parse numbers
    parsed_numbers = [_parse_number(n) for n in numbers]
    
    # Extract product name (text before numbers)
    # Find position of first number
    first_num_match = re.search(r'[\d.,]+', line)
    if first_num_match:
        product = line[:first_num_match.start()].strip()
    else:
        product = line.strip()
    
    # Skip if no product name
    if not product or len(product) < 2:
        return None
    
    # Build item
    item = {
        'product_type': product,
        'denomination': 0,
        'quantity': 0,
        'amount': 0
    }
    
    # Assign numbers (heuristic: last is amount, second-to-last is quantity)
    if len(parsed_numbers) >= 2:
        item['amount'] = parsed_numbers[-1]
        item['quantity'] = parsed_numbers[-2]
    
    if len(parsed_numbers) >= 3:
        item['denomination'] = parsed_numbers[-3]
    
    return item


def extract_totals_from_text_content(text: str) -> Dict[str, float]:
    """
    Extract totals from PDF text.
    """
    totals = {}
    
    lines = text.split('\n')
    
    for line in lines:
        line_lower = line.lower()
        
        # VAT rate
        if 'vat' in line_lower or 'thuế' in line_lower:
            numbers = re.findall(r'[\d.,]+', line)
            for num in numbers:
                val = _parse_number(num)
                if 0 < val < 100 and 'vat_rate' not in totals:
                    totals['vat_rate'] = val
                elif val > 100 and 'vat_amount' not in totals:
                    totals['vat_amount'] = val
        
        # Total before tax
        if any(kw in line_lower for kw in ['tổng trước thuế', 'subtotal', 'before tax']):
            numbers = re.findall(r'[\d.,]+', line)
            for num in numbers:
                val = _parse_number(num)
                if val > 0:
                    totals['total_before_tax'] = val
                    break
        
        # Total payment
        if any(kw in line_lower for kw in ['tổng thanh toán', 'tổng cộng', 'grand total', 'total payment']):
            numbers = re.findall(r'[\d.,]+', line)
            for num in numbers:
                val = _parse_number(num)
                if val > 0:
                    totals['total_payment'] = max(totals.get('total_payment', 0), val)
    
    return totals


def parse_scanned_pdf(pdf_bytes: bytes) -> Dict[str, Any]:
    """
    Parse scanned PDF using OCR.
    Converts PDF pages to images and applies OCR.
    """
    try:
        # Convert PDF to images
        images = convert_from_bytes(pdf_bytes, dpi=300)
        
        all_text_blocks = []
        
        # Process each page
        for page_num, image in enumerate(images):
            # Convert PIL image to numpy array
            image_array = np.array(image)
            
            # Preprocess
            from parsers.invoice_ocr import preprocess_image
            preprocessed = preprocess_image(image_array)
            
            # Run OCR
            ocr = get_ocr_engine()
            result = ocr.ocr(preprocessed, cls=True)
            
            # Extract text blocks
            if result and result[0]:
                for line in result[0]:
                    bbox = line[0]
                    text = line[1][0]
                    confidence = line[1][1]
                    
                    y_pos = bbox[0][1]
                    x_pos = bbox[0][0]
                    
                    all_text_blocks.append({
                        'text': text,
                        'confidence': confidence,
                        'x': x_pos,
                        'y': y_pos + (page_num * 3000),  # Offset for page number
                        'bbox': bbox,
                        'page': page_num
                    })
        
        # Sort by position
        all_text_blocks.sort(key=lambda b: (b['page'], b['y']))
        
        # Extract structured data
        line_items = extract_line_items_from_text(all_text_blocks)
        totals = extract_totals_from_text(all_text_blocks)
        
        return {
            'line_items': line_items,
            'totals': totals
        }
        
    except Exception as e:
        raise Exception(f"Error processing scanned PDF: {str(e)}")


def _parse_number(value) -> float:
    """Parse number from Vietnamese formatted string."""
    if isinstance(value, (int, float)):
        return float(value)
    
    value_str = str(value).strip()
    
    # Remove non-numeric chars except dots, commas, minus
    value_str = re.sub(r'[^0-9.,-]', '', value_str)
    
    # Handle Vietnamese format
    if '.' in value_str and ',' in value_str:
        if value_str.rfind('.') > value_str.rfind(','):
            value_str = value_str.replace('.', '').replace(',', '.')
        else:
            value_str = value_str.replace(',', '')
    elif value_str.count('.') > 1:
        value_str = value_str.replace('.', '')
    elif value_str.count(',') > 1:
        value_str = value_str.replace(',', '')
    
    try:
        return float(value_str) if value_str else 0
    except ValueError:
        return 0
