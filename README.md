# Transparisation
Codebase for automating the transparency process in finance. It integrates data handling, reporting, and monitoring into a streamlined workflow, with an intuitive interface to manage and control each step. Designed to improve efficiency, reduce manual work, and ensure reliability in financial operations.

# Solvency II Transparisation Tool

This project automates the Solvency II transparisation workflow for asset managers and insurers. It ensures regulatory compliance by validating fund data, generating standardized outputs, and streamlining communication.

## Features

- Pre-transparisation comparison of fund portfolios to detect changes
- Cleaning and validation of TPT files with financial metrics (e.g., duration, coupon rate)
- Error reporting and compliance checks
- Generation of SAS-compatible input files
- Automated email dispatch to asset managers with tracking logic

## Technologies

- Python (Pandas, Tkinter, OpenPyXL)
- COM automation via `win32com` for Outlook integration

## Usage

1. Launch the interface and select the relevant TPT files
2. Run pre-transparisation checks to identify portfolio changes
3. Validate and enrich data for regulatory reporting
4. Generate outputs and dispatch emails to asset managers

## License

This project is licensed under the MIT License.
