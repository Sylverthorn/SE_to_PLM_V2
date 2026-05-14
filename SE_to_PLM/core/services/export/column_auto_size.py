import concurrent.futures
from openpyxl.utils import get_column_letter

def calculate_column_width(column_cells) -> float:
    """Calculates the maximum width for a single column's cells."""
    max_length = 0
    for cell in column_cells:
        try:
            if cell.value:
                length = len(str(cell.value))
                if length > max_length:
                    max_length = length
        except:
            continue
    return float(max_length + 2)

def auto_size_columns(worksheet):
    """
    Adjusts the width of all columns in the worksheet based on their content.
    Uses a ThreadPoolExecutor for faster processing of large sheets.
    """
    columns = list(worksheet.columns)
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        # Map each column to a future calculating its width
        future_to_col_idx = {
            executor.submit(calculate_column_width, col): i + 1 
            for i, col in enumerate(columns)
        }
        
        for future in concurrent.futures.as_completed(future_to_col_idx):
            col_idx = future_to_col_idx[future]
            width = future.result()
            column_letter = get_column_letter(col_idx)
            # Limit maximum width to 100 for sanity
            worksheet.column_dimensions[column_letter].width = min(width, 100)
