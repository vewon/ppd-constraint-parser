# PPD Constraint Parser

## Overview
This Python script parses **PostScript Printer Description (PPD)** files to extract option constraints and creates a pivot table showing the relationships between the printer options. It helps visualize which combinations of options are restricted by the printer driver.

The script does the following:
- Reads the PPD file
- Extracts the **UIConstraints**, **OpenUI/CloseUI blocks**, and option labels
- Creates a pivot table for a chosen option, showing how it interact with other related options
- Outputs results to the console, CSV, and Excel formats.

## Features
- **PPD Parsing**: Extracts the constraints, options, and labels
- **Pivot Table Generation**: Display relationships between a target option and related option

## Requirements
- `Python`: 3.11+
- Dependencies:
    - `pandas`: 3.0.1

To install the dependencies, use the following command:
`pip install pandas`

## Usage
Run the script from the command line using the following:
`python constraint_analyzer.py <ppd_file> <target_option>`

## Output

The pivot table includes the following:
1. RelatedOption: The option constrained with the target option
2. RelatedValue: The specific value of the related option
3. Columns: Values of the target option
4. Cells: "X" marks the constrained combinations

| RelatedOption | RelatedValue | None | DuplexNoTumble | DuplexTumble |
|---------------|--------------|------|----------------|--------------|
|               |              | Off  | Long Edge      | Short Edge   |
| MediaType     | Plain        |      |                |              |
| MediaType     | Bond         |      |                |              |
| MediaType     | Cardstock    |      |                |              |
| MediaType     | Glossy       |      | X              | X            |
| MediaType     | Labels       |      | X              | X            |