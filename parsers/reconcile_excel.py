"""
Reconciliation Excel Parser

Parses reconciliation Excel files with Vietnamese headers.
Handles multiple sections and flexible column layouts.
"""

import pandas as pd
import re
from typing import Dict, List, Any


def parse_reconciliation_excel(file) -> Dict[str, Any]:
    """
    Parse reconciliation Excel file.
    
    Args:
        file: File-like object or path to Excel file
        
    Returns:
        Dictionary with 'line_items' DataFrame and 'totals' dict
    """
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(file)
        all_line_items = []
        totals = {}
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file, sheet_name=sheet_name, header=None)
            
            # Extract line items from this sheet
            items, sheet_totals = _extract_line_items_from_sheet(df)
            all_line_items.extend(items)
            
            # Merge totals (last sheet wins)
            if sheet_totals:
                totals.update(sheet_totals)
        
        # Convert to DataFrame
        if all_line_items:
            line_items_df = pd.DataFrame(all_line_items)
        else:
            # Return empty DataFrame with expected columns
            line_items_df = pd.DataFrame(columns=[
                'product_type', 'denomination', 'quantity', 'amount'
            ])
        
        return {
            'line_items': line_items_df,
            'totals': totals
        }
        
    except Exception as e:
        raise Exception(f"Error parsing reconciliation Excel: {str(e)}")


def _extract_line_items_from_sheet(df: pd.DataFrame) -> tuple:
    """
    Extract line items from a single sheet.
    Handles multiple table sections within the sheet.
    
    Returns:
        Tuple of (line_items_list, totals_dict)
    """
    line_items = []
    totals = {}
    
    # Find all header rows (rows containing key Vietnamese terms)
    header_patterns = [
        r'loại\s*sản\s*phẩm',
        r'mệnh\s*giá',
        r'số\s*lượng',
        r'thành\s*tiền'
    ]
    
    potential_headers = []
    for idx, row in df.iterrows():
        row_str = ' '.join([str(cell).lower() for cell in row if pd.notna(cell)])
        if any(re.search(pattern, row_str) for pattern in header_patterns):
            potential_headers.append(idx)
    
    # Extract data from each header section
    for header_idx in potential_headers:
        items = _extract_section(df, header_idx)
        line_items.extend(items)
    
    # Extract totals (look for VAT, total payment keywords)
    totals = _extract_totals(df)
    
    return line_items, totals


def _extract_section(df: pd.DataFrame, header_idx: int) -> List[Dict]:
    """
    Extract line items from a section starting at header_idx.
    """
    items = []
    
    # Map columns
    header_row = df.iloc[header_idx]
    column_map = _map_columns(header_row)
    
    if not column_map:
        return items
    
    # Extract data rows (until empty row or another header)
    for idx in range(header_idx + 1, len(df)):
        row = df.iloc[idx]
        
        # Stop at empty row or summary keywords
        if _is_empty_row(row) or _is_summary_row(row):
            break
        
        # Extract item
        item = _extract_item(row, column_map)
        if item:
            items.append(item)
    
    return items


def _map_columns(header_row: pd.Series) -> Dict[str, int]:
    """
    Map Vietnamese column headers to column indices.
    """
    column_map = {}
    
    for idx, cell in enumerate(header_row):
        cell_str = str(cell).lower().strip()
        
        # Product type
        if re.search(r'loại|sản\s*phẩm|tên', cell_str):
            column_map['product'] = idx
        
        # Denomination
        elif re.search(r'mệnh\s*giá|denomination', cell_str):
            column_map['denomination'] = idx
        
        # Quantity
        elif re.search(r'số\s*lượng|quantity', cell_str):
            column_map['quantity'] = idx
        
        # Amount
        elif re.search(r'thành\s*tiền|amount|tổng', cell_str) and 'tổng cộng' not in cell_str:
            column_map['amount'] = idx
        
        # Discount (optional)
        elif re.search(r'chiết\s*khấu|discount|giảm\s*giá', cell_str):
            column_map['discount'] = idx
    
    return column_map


def _extract_item(row: pd.Series, column_map: Dict[str, int]) -> Dict:
    """
    Extract a single line item from a row.
    """
    try:
        product = str(row[column_map['product']]) if 'product' in column_map else ''
        denomination = row[column_map['denomination']] if 'denomination' in column_map else 0
        quantity = row[column_map['quantity']] if 'quantity' in column_map else 0
        amount = row[column_map['amount']] if 'amount' in column_map else 0
        
        # Skip if missing critical data
        if pd.isna(product) or product.strip() == '' or product == 'nan':
            return None
        
        # Parse denomination (handle various formats)
        if pd.notna(denomination):
            denomination = _parse_number(denomination)
        else:
            denomination = 0
        
        # Parse quantity
        if pd.notna(quantity):
            quantity = _parse_number(quantity)
        else:
            quantity = 0
        
        # Parse amount
        if pd.notna(amount):
            amount = _parse_money(amount)
        else:
            amount = 0
        
        # Extract discount if available
        discount = 0
        if 'discount' in column_map:
            discount_val = row[column_map['discount']]
            if pd.notna(discount_val):
                discount = _parse_money(discount_val)
        
        return {
            'product_type': product.strip(),
            'denomination': denomination,
            'quantity': quantity,
            'amount': amount,
            'discount': discount
        }
        
    except Exception:
        return None


def _extract_totals(df: pd.DataFrame) -> Dict[str, float]:
    """
    Extract summary totals (VAT, total payment) from sheet.
    """
    totals = {}
    
    for idx, row in df.iterrows():
        row_str = ' '.join([str(cell).lower() for cell in row if pd.notna(cell)])
        
        # Look for VAT rate
        if re.search(r'thuế\s*vat|vat\s*rate|%\s*vat', row_str):
            for cell in row:
                if pd.notna(cell) and isinstance(cell, (int, float)):
                    if 0 < cell < 100:  # Likely a percentage
                        totals['vat_rate'] = cell
                        break
        
        # Look for VAT amount
        if re.search(r'tiền\s*thuế|vat\s*amount', row_str):
            for cell in row:
                if pd.notna(cell) and isinstance(cell, (int, float, str)):
                    val = _parse_money(cell)
                    if val > 0:
                        totals['vat_amount'] = val
                        break
        
        # Look for total before tax
        if re.search(r'tổng\s*trước\s*thuế|before\s*tax|subtotal', row_str):
            for cell in row:
                if pd.notna(cell) and isinstance(cell, (int, float, str)):
                    val = _parse_money(cell)
                    if val > 0:
                        totals['total_before_tax'] = val
                        break
        
        # Look for total payment
        if re.search(r'tổng\s*thanh\s*toán|total\s*payment|grand\s*total|tổng\s*cộng', row_str):
            for cell in row:
                if pd.notna(cell) and isinstance(cell, (int, float, str)):
                    val = _parse_money(cell)
                    if val > 0:
                        totals['total_payment'] = val
                        break
    
    return totals


def _parse_number(value) -> float:
    """
    Parse number from various formats.
    Handles Vietnamese thousand separators (.) and decimal (,)
    """
    if isinstance(value, (int, float)):
        return float(value)
    
    value_str = str(value).strip()
    
    # Remove currency symbols and whitespace
    value_str = re.sub(r'[₫đvnd\s]', '', value_str, flags=re.IGNORECASE)
    
    # Handle Vietnamese format: 1.000.000,50 or 1,000,000.50
    # Count dots and commas to determine format
    dot_count = value_str.count('.')
    comma_count = value_str.count(',')
    
    if dot_count > 1 or (dot_count == 1 and comma_count == 1 and value_str.rfind('.') > value_str.rfind(',')):
        # Vietnamese format: dots for thousands, comma for decimal
        value_str = value_str.replace('.', '').replace(',', '.')
    elif comma_count > 1:
        # US format with comma thousands
        value_str = value_str.replace(',', '')
    
    # Extract number
    match = re.search(r'-?[\d.]+', value_str)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return 0
    
    return 0


def _parse_money(value) -> float:
    """
    Parse money value (same as _parse_number but more explicit for amounts)
    """
    return _parse_number(value)


def _is_empty_row(row: pd.Series) -> bool:
    """Check if row is empty or mostly empty."""
    non_null = row.notna().sum()
    return non_null <= 1


def _is_summary_row(row: pd.Series) -> bool:
    """Check if row contains summary keywords."""
    row_str = ' '.join([str(cell).lower() for cell in row if pd.notna(cell)])
    summary_patterns = [
        r'tổng\s*cộng',
        r'tổng\s*thanh\s*toán',
        r'thuế\s*vat',
        r'grand\s*total',
        r'total\s*payment'
    ]
    return any(re.search(pattern, row_str) for pattern in summary_patterns)
