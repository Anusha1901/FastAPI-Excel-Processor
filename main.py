from fastapi import FastAPI, HTTPException, Query
from typing import List, Dict, Any, Optional
import pandas as pd
import numpy as np
from pathlib import Path
import logging
from dataclasses import dataclass
import re

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel Processor API",
    description="FastAPI application for processing Excel files with multiple tables",
    version="1.0.0"
)

@dataclass
class TableInfo:
    """Data class to store table information"""
    name: str
    start_row: int
    end_row: int
    start_col: int
    end_col: int
    data: pd.DataFrame

class ExcelProcessor:
    """Class to handle Excel file processing and table extraction"""
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.tables: Dict[str, TableInfo] = {}
        self.raw_data: Optional[pd.DataFrame] = None
        self._load_and_parse_excel()
    
    def _load_and_parse_excel(self) -> None:
        """Load Excel file and parse tables"""
        try:
            if not self.file_path.exists():
                raise FileNotFoundError(f"Excel file not found: {self.file_path}")
            
            # Read the Excel file
            self.raw_data = pd.read_excel(
                self.file_path, 
                sheet_name=0,  # First sheet
                header=None,   # Don't treat first row as header
                engine='xlrd' if self.file_path.suffix == '.xls' else 'openpyxl'
            )
            
            logger.info(f"Loaded Excel file with shape: {self.raw_data.shape}")
            
            # Print first few rows for debugging
            logger.info("First 10 rows of data:")
            for i in range(min(10, len(self.raw_data))):
                row_data = [str(val) if pd.notna(val) else 'NaN' for val in self.raw_data.iloc[i, :8]]
                logger.info(f"Row {i}: {row_data}")
            
            self._identify_tables_improved()
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Error loading Excel file: {str(e)}")
    
    def _identify_tables_improved(self) -> None:
        """Improved table identification for capital budgeting Excel files"""
        if self.raw_data is None:
            return
        
        # Define known table patterns for capital budgeting files
        table_patterns = self.raw_data
        
        # Find sections based on key indicators
        for table_name, keywords in table_patterns.items():
            table_info = self._find_table_by_keywords(table_name, keywords)
            if table_info:
                self.tables[table_name] = table_info
        
        # Also try to identify tables by looking for section headers
        self._find_section_headers()
        
        logger.info(f"Identified {len(self.tables)} tables: {list(self.tables.keys())}")
    
    def _find_table_by_keywords(self, table_name: str, keywords: List[str]) -> Optional[TableInfo]:
        """Find table based on keyword patterns"""
        try:
            # Convert data to string for searching
            data_str = self.raw_data.astype(str).fillna('')
            
            # Find rows that contain the keywords
            matching_rows = []
            for i in range(len(data_str)):
                row_text = ' '.join(data_str.iloc[i, :].values).lower()
                keyword_matches = sum(1 for keyword in keywords if keyword.lower() in row_text)
                if keyword_matches >= 1:  # At least one keyword match
                    matching_rows.append(i)
            
            if not matching_rows:
                return None
            
            # Define table boundaries
            start_row = min(matching_rows)
            end_row = max(matching_rows)  # Add some buffer
            
            # Find the actual data columns
            start_col = 0
            end_col = min(7, len(self.raw_data.columns))  # Limit to first 10 columns
            
            
            # Extract table data
            table_data = self.raw_data.iloc[start_row:end_row, start_col:end_col].copy()
            
            return TableInfo(
                name=table_name,
                start_row=start_row,
                end_row=end_row,
                start_col=start_col,
                end_col=end_col,
                data=table_data
            )
            
        except Exception as e:
            logger.warning(f"Error finding table {table_name}: {str(e)}")
            return None
    
    def _find_section_headers(self) -> None:
        """Find section headers in the Excel file"""
        if self.raw_data is None:
            return
        
        data_str = self.raw_data.astype(str).fillna('')
        
        # Look for cells that look like section headers
        section_headers = []
        for i in range(len(data_str)):
            for j in range(min(3, len(data_str.columns))):  # Check first 3 columns
                cell_value = data_str.iloc[i, j].strip()
                if (len(cell_value) > 5 and 
                    cell_value.isupper() and 
                    not cell_value.replace(' ', '').replace('=', '').isdigit()):
                    section_headers.append((cell_value, i, j))
        
        # Create tables for each section header
        for header, row, col in section_headers[:5]:  # Limit to first 5 headers
            if header not in self.tables:
                table_info = self._extract_section_table(header, row, col)
                if table_info:
                    self.tables[header] = table_info
    
    def _extract_section_table(self, header: str, start_row: int, start_col: int) -> Optional[TableInfo]:
        """Extract table data for a section"""
        try:
            # Define reasonable boundaries
            end_row = min(start_row + 20, len(self.raw_data))
            end_col = min(start_col + 8, len(self.raw_data.columns))
            
            # Extract data
            table_data = self.raw_data.iloc[start_row:end_row, start_col:end_col].copy()
            
            return TableInfo(
                name=header,
                start_row=start_row,
                end_row=end_row,
                start_col=start_col,
                end_col=end_col,
                data=table_data
            )
            
        except Exception as e:
            logger.warning(f"Error extracting section table {header}: {str(e)}")
            return None
    
    def get_table_names(self) -> List[str]:
        """Get list of all table names"""
        return list(self.tables.keys())
    
    def get_table_row_names(self, table_name: str) -> List[str]:
        """Get row names (first column values) for a specific table"""
        if table_name not in self.tables:
            raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found")
        
        table_info = self.tables[table_name]
        
        # Get all values from first column
        first_column = table_info.data.iloc[:, 0]
        
        # Filter and clean row names
        row_names = []
        for val in first_column:
            if pd.notna(val):
                str_val = str(val).strip()
                if (str_val and 
                    str_val.lower() != 'nan' and 
                    len(str_val) > 1 and
                    not str_val.replace('.', '').replace('-', '').isdigit()):
                    row_names.append(str_val)
        
        return row_names
    
    def calculate_row_sum(self, table_name: str, row_name: str) -> float:
        """Calculate sum of numerical values in a specific row"""
        if table_name not in self.tables:
            raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found")
        
        table_info = self.tables[table_name]
        
        # Find the row with the specified name (flexible matching)
        target_row_idx = None
        for idx, val in enumerate(table_info.data.iloc[:, 0]):
            if pd.notna(val):
                cell_value = str(val).strip()
                # Try exact match first, then partial match
                if (cell_value == row_name or 
                    row_name in cell_value or 
                    cell_value in row_name):
                    target_row_idx = idx
                    break
        
        if target_row_idx is None:
            # Try searching in all columns for the row name
            for idx in range(len(table_info.data)):
                row_data = table_info.data.iloc[idx, :]
                for val in row_data:
                    if pd.notna(val) and row_name.lower() in str(val).lower():
                        target_row_idx = idx
                        break
                if target_row_idx is not None:
                    break
        
        if target_row_idx is None:
            available_rows = self.get_table_row_names(table_name)
            raise HTTPException(
                status_code=404, 
                detail=f"Row '{row_name}' not found in table '{table_name}'. Available rows: {available_rows}"
            )
        
        # Get the row data (all columns)
        row_data = table_info.data.iloc[target_row_idx, :]
        
        # Calculate sum of numerical values
        total_sum = 0
        values_found = []
        
        for val in row_data:
            if pd.notna(val):
                try:
                    str_val = str(val).strip()
                    if str_val and str_val.lower() != 'nan':
                        # Handle percentage values
                        if str_val.endswith('%'):
                            num_val = int(str_val[:-1])
                        else:
                            # Try to extract number from string
                            num_str = re.sub(r'[^\d.-]', '', str_val)
                            if num_str and num_str != '-':
                                num_val = int(num_str)
                            else:
                                continue
                        
                        total_sum += num_val
                        values_found.append(num_val)
                        
                except (ValueError, TypeError):
                    continue
        
        logger.info(f"Row '{row_name}' in table '{table_name}': found values {values_found}, sum = {total_sum}")
        return total_sum

# Initialize the Excel processor
excel_processor = None

@app.on_event("startup")
async def startup_event():
    """Initialize the Excel processor on startup"""
    global excel_processor
    try:
        excel_file_path = "Data/capbudg.xls"
        excel_processor = ExcelProcessor(excel_file_path)
        logger.info("Excel processor initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize Excel processor: {str(e)}")
        raise

@app.get("/")
async def root():
    """Root endpoint with API information"""
    return {
        "message": "Excel Processor API",
        "version": "1.0.0",
        "endpoints": [
            "/list_tables",
            "/get_table_details",
            "/row_sum"
        ]
    }

@app.get("/list_tables")
async def list_tables():
    """
    List all table names present in the Excel sheet.
    
    Returns:
        dict: Dictionary containing list of table names
    """
    try:
        if excel_processor is None:
            raise HTTPException(status_code=500, detail="Excel processor not initialized")
        
        table_names = excel_processor.get_table_names()
        return {"tables": table_names}
    
    except Exception as e:
        logger.error(f"Error in list_tables: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/get_table_details")
async def get_table_details(table_name: str = Query(..., description="Name of the table")):
    """
    Get row names (first column values) for a specific table.
    
    Args:
        table_name: Name of the table to get details for
        
    Returns:
        dict: Dictionary containing table name and row names
    """
    try:
        if excel_processor is None:
            raise HTTPException(status_code=500, detail="Excel processor not initialized")
        
        row_names = excel_processor.get_table_row_names(table_name)
        return {
            "table_name": table_name,
            "row_names": row_names
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in get_table_details: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/row_sum")
async def row_sum(
    table_name: str = Query(..., description="Name of the table"),
    row_name: str = Query(..., description="Name of the row")
):
    """
    Calculate the sum of all numerical values in a specific row of a table.
    
    Args:
        table_name: Name of the table
        row_name: Name of the row (must be one from get_table_details)
        
    Returns:
        dict: Dictionary containing table name, row name, and calculated sum
    """
    try:
        if excel_processor is None:
            raise HTTPException(status_code=500, detail="Excel processor not initialized")
        
        calculated_sum = excel_processor.calculate_row_sum(table_name, row_name)
        return {
            "table_name": table_name,
            "row_name": row_name,
            "sum": calculated_sum
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in row_sum: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

# Add a debug endpoint to help understand the Excel structure
@app.get("/debug_excel")
async def debug_excel():
    """
    Debug endpoint to see the raw Excel data structure
    """
    try:
        if excel_processor is None:
            raise HTTPException(status_code=500, detail="Excel processor not initialized")
        
        # Get first 20 rows and 10 columns for debugging
        debug_data = []
        for i in range(min(20, len(excel_processor.raw_data))):
            row_data = []
            for j in range(min(10, len(excel_processor.raw_data.columns))):
                val = excel_processor.raw_data.iloc[i, j]
                row_data.append(str(val) if pd.notna(val) else "")
            debug_data.append({f"row_{i}": row_data})
        
        return {
            "excel_shape": excel_processor.raw_data.shape,
            "first_20_rows": debug_data,
            "detected_tables": list(excel_processor.tables.keys())
        }
    
    except Exception as e:
        logger.error(f"Error in debug_excel: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=9090)
