# Competitive Pricing Analysis System

This README provides a brief overview of the project.  

Detailed documentation, input requirements, pricing logic, report descriptions, and usage instructions are available in the **User Guide** file.


## Overview

The Competitive Pricing Analysis System is a Python-based business analytics tool designed to evaluate a company’s pricing position relative to multiple competitors.

The system processes company and competitor pricing datasets, compares prices across comparable products, and generates automated Excel reports containing pricing insights, charts, and strategic observations.

The analysis is fully automated and requires only:
-	a ZIP archive containing pricing datasets;
-	selection of the company pricing file;
-	a user-defined pricing threshold.
-	an output Excel filename.


## Business Value

The system helps organisations:

-	identify underpriced, overpriced, market-aligned, and exclusive products;
-	improve pricing transparency across the product assortment;
-	detect potential pricing risks and inefficiencies;
-	evaluate pricing patterns within individual product categories;
-	support data-driven pricing decisions through automated analytical insights and recommendations.


## Key Features

-	Pricing Category Classification (A–K)  
Automatically classifies products into pricing categories using a configurable pricing threshold.

-	Pricing Deviation Analysis  
Calculates pricing deviations between company and competitor prices to quantify pricing differences.

-	Item Category Analysis  
Identifies pricing patterns and distribution trends within individual product categories.

-	Pricing Structure & Risk Assessment  
Evaluates overall assortment pricing patterns and detects potential pricing risks.

-	Automated Strategic Recommendations  
Generates data-driven recommendations based on pricing behaviour and category distributions.


## Input Files

The system requires a ZIP archive containing:

- 1 company pricing file (.xlsx)
- 3 competitor pricing files (.xlsx)


## Output Files

The system generates a multi-sheet Excel report containing:

- Pricing Analysis Summary
- Pricing Analysis Data
- Item Category Analysis
- Filtered Pricing Category Worksheets

The report includes pricing classifications, pricing deviations, distribution charts, pricing structure assessment, risk evaluation, and automated recommendations.


## Software Requirements

-	Python 3.9 or later (or a compatible Python execution environment).
-	Required libraries: pandas and openpyxl


## Usage

1.	Prepare a ZIP archive containing the required price files.
2.	Run the program.
3.	Enter the path to the ZIP archive. 
4.	Select the company pricing file.
5.	Enter the pricing threshold percentage.
6.	Specify the output Excel filename.
7.	Review the generated pricing analysis report.


## License

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

This project is licensed under the MIT License.
See the **LICENSE** file for details.



## Author

Polina Zimina
