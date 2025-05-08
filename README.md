# Data Curation Workflows
authors: Ellery Galvin
date: 2024-04-29
contact: ellery.galvin@colorado.edu


## Directory contents
- detectors.py implements a series of automated tests for data curation workflows.
- data contains spreadsheets on which to test detectors.py
- .gitignore excludes the data directory from version control
- environment.yml contains the conda environment used to run the tests


## Tests implemented to date
- filename short enough
- table location in the upper left corner
- column headers acceptable
- no special characters
- no untrimmed white space
- no question mark only cells
- no whitespace only cells
- dates use ISO format


## Usage
- Create a conda environment from the environment.yml file
```bash
conda env create -f environment.yml
```
- Activate the conda environment
```bash
conda activate excel
```
- Run the main method
```bash
ipython detectors.py 
```
Python will output the results of the tests on the demo.xlsx file in the data directory and save the results in both a json and csv file.
