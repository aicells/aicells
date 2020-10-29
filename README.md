# AIcells

AIcells is an Excel add-in that lets you work with Python functions in Excel.

Statistics, Bayesian theory, Machine Learning, and Optimization in Excel cells. Many useful analytical methods and Python functions.

AIcells is an open-source licensed Python library that makes it easy to make analysis in Excel and even continue in Python.

The project is currently in early alpha stage.


## System requirements

AIcells requires a 64bit Excel from Office 365 on Windows.

The Excel version has to support the [UNIQUE](https://support.office.com/en-us/article/unique-function-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e) function.

## Installation

You can download the latest pre-release (aicells-{version}.zip) from [here.](https://github.com/aicells/aicells/releases)  

Unzip the release zip file. It contains a root folder called "aicells-{version}".

To use AIcells' UDFs from Excel, you have to:
*	Install AICells' Excel add-in
*	And run our UDF server

### How to install AIcells' Excel Add-in

1.	On Excel's "File" tab click on "Options" and choose the "Add-Ins" category
2.	Select "Excel Add-ins" from the drop-down box at the bottom of the dialog, then click on the "Go" button
3.	Click on the "Browse" button and choose aicells-{version}/aicells-excel-add-in/aicells.xlam and click on the "OK" button

You can find additional information in Microsoft's article: [Add or remove add-ins in Excel](https://support.office.com/en-us/article/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460)

For more information about add-in installation, visit AutomateExcel.com's article:  [How to Install (or Uninstall) a VBA add-in (.xlam file) for Microsoft Excel](https://www.automateexcel.com/vba/install-add-in)

### Running the UDF server

Run aicells-{version}/aicells.bat . When it shows the message "xlwings server running, clsid={...}", 
the UDF server is running, you can use AIcells' UDFs from Excel.

### Examples

Check out our examples from the aicells-{version}/aicells-examples/ directory.

### Technology Stack

* We use [WinPython](https://winpython.github.io/) portable portable distribution to execute the UDF server written in python
* We use [xlwings](https://github.com/xlwings/xlwings) for the Excel-Python communication
 
### Plugins

We have five plugins in AICells version 0.0.2:
* Core: Function and tool parameter lists, CSV read.write support
* Correlation: Correlation matrix calculation
* Random: Random number generators
* Seaborn: Seaborn charting support, it can insert SVG Seaborn charts into the worksheet
* Supervised-learning: Regression models from scikit-learn and XGBoost
 
