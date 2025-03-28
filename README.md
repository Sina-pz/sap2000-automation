# SAP2000 Automation with Python

This project demonstrates how to automate SAP2000 structural analysis using Python with COM (Component Object Model) interface.

## Tutorial Example

The script automates the analysis of a simply supported beam with the following properties:

- Beam Length: 6 meters
- Cross-section: Rectangular beam (300 mm × 500 mm)
- Material: Concrete (fc' = 25 MPa)
- Load Type: Uniformly distributed load (UDL)
- Load Magnitude: 5 kN/m
- Supports: Simple supports at both ends (pinned at left, roller at right)

## Prerequisites

1. SAP2000 installed on your Windows machine
2. Python 3.6 or higher
3. Required Python packages (install using `pip install -r requirements.txt`):
   - pywin32
   - comtypes

## Setup Instructions

1. Make sure SAP2000 is properly installed and licensed on your system
2. Create a virtual environment (recommended):
   ```
   python -m venv venv
   ```
3. Activate the virtual environment:
   ```
   .\venv\Scripts\Activate.ps1  # For PowerShell
   ```
   or
   ```
   .\venv\Scripts\activate.bat  # For Command Prompt
   ```
4. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Running the Script

1. Make sure SAP2000 is closed (or the script will connect to the running instance)
2. Run the script:
   ```
   python app.py
   ```
3. The script will:
   - Launch SAP2000
   - Create a new model
   - Define materials and section properties
   - Create a 6m beam
   - Apply a 5 kN/m distributed load
   - Set boundary conditions (pinned and roller supports)
   - Run the analysis
   - Extract and display the results
   - Save the model as "BeamModel.sdb" in the current directory

## Expected Results

The script will print the analysis results and compare them with theoretical calculations:

- Support Reactions: 15 kN at each end
- Maximum Bending Moment: 22.5 kN.m at midspan
- Maximum Shear Force: 15 kN at supports

## Troubleshooting

If you encounter any issues:

1. Make sure SAP2000 is properly installed and licensed
2. Check that the COM interface is working (try running SAP2000 manually first)
3. Verify that the pywin32 and comtypes packages are correctly installed
4. Check the error messages in the console output for specific issues

## Extending the Script

You can modify the script to:
- Change beam properties (length, cross-section, material)
- Apply different load types or magnitudes
- Change support conditions
- Extract additional results
- Create more complex structural models 