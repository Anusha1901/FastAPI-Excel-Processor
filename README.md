# FastAPI Excel Processor

A FastAPI application that processes capital budgeting Excel files and provides REST API endpoints to extract financial data from multiple tables.

## Installation

```bash
pip install -r requirements.txt
```

## Setup

1. Create `Data` folder in project root
2. Place your `capbudg.xls` file in the `Data` folder
3. Run the application:

```bash
python run.py
```

## API Endpoints

### Base URL: `http://localhost:9090`

### 1. GET `/`
Returns API information and available endpoints.

### 2. GET `/list_tables`
Lists all detected table names from the Excel file.

**Response:**
```json
{
  "tables": ["Initial Investment", "DISCOUNT RATE", "WORKING CAPITAL"]
}
```

### 3. GET `/get_table_details?table_name={name}`
Returns row names for a specific table.

**Parameters:**
- `table_name`: Name of the table (from list_tables response)

**Response:**
```json
{
  "table_name": "Initial Investment",
  "row_names": ["Initial Investment=", "Tax Credit (if any )=", "Salvage Value at end of project="]
}
```

### 4. GET `/row_sum?table_name={name}&row_name={row}`
Calculates sum of numerical values in a specific row.

**Parameters:**
- `table_name`: Name of the table
- `row_name`: Name of the row (from get_table_details response)

**Response:**
```json
{
  "table_name": "Initial Investment",
  "row_name": "Tax Credit (if any )=",
  "sum": 10
}
```

### 5. GET `/debug`
Shows raw Excel data structure for debugging.

## Usage Example

1. List tables: `http://localhost:9090/list_tables`
2. Get table details: `http://localhost:9090/get_table_details?table_name=Initial%20Investment`
3. Calculate row sum: `http://localhost:9090/row_sum?table_name=Initial%20Investment&row_name=Tax%20Credit%20(if%20any%20)%3D`

## API Documentation

- Swagger UI: `http://localhost:9090/docs`
- ReDoc: `http://localhost:9090/redoc`

## File Structure

```
project/
├── main.py              # FastAPI application
├── run.py               # Application runner
├── requirements.txt     # Dependencies
├── README.md           # This file
└── Data/
    └── capbudg.xls     # Excel file to process
```

## Dependencies

- FastAPI
- pandas
- openpyxl
- xlrd
- uvicorn

## My Insights

### Potential Improvements

The application could be enhanced by supporting multiple Excel file formats (.xlsx, .csv) and implementing dynamic file upload functionality instead of requiring files in a specific directory.  

### Missed Edge Cases

The current implementation may not handle completely empty Excel files or files with only header rows without any numerical data, which could cause processing errors.  The application assumes a specific Excel structure and may fail with files that have merged cells, multiple sheets, or unconventional layouts. 
