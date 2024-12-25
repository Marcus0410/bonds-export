# bonds-export

`bonds-export` is a Go application that processes Excel files containing trade allocations and generates two output Excel files. The output files contain the data from the input file but formatted differently for further use.

## Features

- **Input Processing:** The script reads an Excel file containing trade allocation details.
- **Data Transformation:** It formats and structures the data into two separate output files.
- **Output Files:** Generates two `.xlsx` files:
  - **Trade Upload File:** Contains a formatted list of allocations with specific columns.
  - **Finance Report File:** Contains a finance report with trade details and calculations.

## Requirements

- **Go 1.18+**
- **Excelize Library**: The Go `excelize` package is used for working with Excel files.

Install the required Go package:
```bash
go get github.com/xuri/excelize/v2
```

## How to Run

1. Clone the repository:

```bash
git clone https://github.com/yourusername/bonds-export.git
cd bonds-export
```

2. Place your input Excel file in the `input` directory. The file should be an `.xlsx` file containing trade allocation data.
3. Run the Go application:

```bash
go run main.go
```

4. The script will process the input file and generate the following output files in the `output` directory:
   - `tradeUpload.xlsx`: Contains formatted trade allocation data.
   - `finance.xlsx`: Contains the finance report with the necessary calculations.

## Directory Structure

```
bonds-export/
│
├── input/              # Input directory for Excel files
│   └── your-input.xlsx
│
├── output/             # Output directory for generated files
│   ├── tradeUpload.xlsx
│   └── finance.xlsx
│
├── main.go             # Main Go application file
└── README.md           # Project documentation
```

## Explanation of Key Functions

- `getInputFilePath`: Finds the most recently modified `.xlsx` file in the `input` directory.
- `readInput`: Reads and processes data from the input Excel file and maps it into Go structures.
- `writeTradeUpload`: Creates the trade upload file (`tradeUpload.xlsx`) from the processed data.
- `writeFinance`: Generates the finance report file (`finance.xlsx`) with calculated values based on allocations.
- `filteredBandD`: Filters the allocations to include only those with `ABG` as the `bAndD` value.
