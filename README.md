# Arrow Ship Mode & OH Merge Automation

Python automation that combines weekly On-Hand and Arrow Ship Mode Excel data into a consolidated Purchasing report.

## Overview

This project automates a recurring Excel-based reporting workflow used by a Purchasing team.

The automation reads two Excel datasets, prepares the data, applies business rules, merges the records using Customer Part Number, filters the required Purchasing Group records, and generates a consolidated Excel report.

The automation is used in a Purchasing team production workflow.

## Business Problem

The weekly process required repetitive manual Excel preparation, including filling down repeated values, maintaining invoice information, matching datasets, filtering records, and preparing the final report.

The automation standardizes these steps using Python and pandas.

## Solution

The script:

1. Determines the current reporting week.
2. Locates the required On-Hand and Arrow Ship Mode files using configured paths.
3. Reads and prepares both datasets.
4. Forward-fills selected Arrow Ship Mode fields.
5. Propagates Invoice Number when consecutive records share the same Customer PO.
6. Normalizes Customer Part Number values.
7. Merges the datasets using Customer Part Number.
8. Filters records where `PGr` starts with `U`.
9. Selects the required reporting columns.
10. Generates the consolidated Excel report.

## Workflow

The detailed processing flow is available in ![docs/diagrams/workflow.png](docs/diagrams/workflow.png)

## Architecture

The application uses a simple Python/pandas script-based architecture. The architecture diagram is available in ![docs/diagrams/architecture.png](docs/diagrams/architecture.png)

## Technologies

* Python
* pandas
* openpyxl
* Microsoft Excel

## Key Features

* Weekly file-path handling
* Excel data processing
* Forward-fill transformation
* Invoice number propagation
* Customer Part Number matching
* DataFrame merge
* Purchasing Group filtering
* Excel report generation

## Input & Output

**Input**

* VN01 On-Hand Part Excel data
* Arrow Ship Mode Excel data

The `PGr` used for filtering comes from the On-Hand dataset.

**Output**

A consolidated Excel report containing the required Purchasing and shipment fields.

## Results

The automation standardizes a recurring weekly reporting process and reduces repetitive manual data preparation.

No percentage or time-saving metric is stated because verified measurements are not available.

## My Role

I developed and maintained the automation, translating Purchasing requirements into Python/pandas data-processing rules and maintaining the reporting workflow.

## Limitations

* Depends on the expected structure and column names of the source Excel files.
* File locations are environment-specific.
* The merge relationship is not explicitly validated.
* The final Arrow Ship Mode row is excluded before processing.
* Customer Part Number normalization could be made more robust.
* Error handling is currently basic.

## Future Improvements

* Input file and column validation
* Safer Customer Part Number normalization
* More explicit handling of non-detail rows
* More detailed processing feedback

## Disclaimer

This repository contains a sanitized portfolio version. Company-specific files, confidential business data, and internal network locations are excluded.
