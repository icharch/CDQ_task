
# CDQ Recruitment task

## Project Description
The aim of the project is to analyse the provided data set to ensure data quality by checking the completeness and correctness of the values in each column, highlighting outliers and proposing corrective actions. Once the data has been corrected, visualisations should be created to answer the following business questions: 
1. What is the trend in the average age of cars over time?
2. Is there a shift in customer preference from petrol to electric vehicles after the Green Deal article? 
3. Does the data show a preference for German car brands among customers?

The project includes two .xlsx files:
1. Task.xlsx: This file contains the input data on which the script will work. The data in Task.xlsx has been transformed into a table using Power Query, and empty rows have been removed.
2. Task_output.xlsx: This file contains the output data obtained after running the script, along with conclusions drawn after executing the script.

## Table of Contents
1. [Installation](#installation)
2. [Usage](#usage)

## Installation
To get started with this project, you need to have Python installed on your system. Then, you can clone this repository and install the necessary dependencies.

### Step-by-Step Installation Guide
1. Clone the repository:
   ```bash
   git clone git@github.com:icharch/CDQ_task.git
   cd cdq_task
### Create a virtual environment
On Windows:
`venv\Scripts\activate`
On macOS and Linux:
`source venv/bin/activate`

### Installing Required Packages

Once your virtual environment is activated, install the necessary packages:
```
pip install pandas plotly openpyxl numpy
```
### Usage
In order to run the script execute
``` 
pyhton task1.py
