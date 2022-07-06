###**wagnerlab_data_processing**

This project consists of python scripts written for data processing and analysis for the wagner lab at GWU.
___
###Table of Contents:
- [1. Requirements](#1-requirements)
- [2. Installation](#2-installation)
- [3. Battery Data Processing](#3-battery-data-processing)
  - [3.1. Classes](#31-classes)
  - [3.2. Functions](#32-functions)
___
#### 1. Requirements
Python installation = 3.8 or later

Required packages: pandas, matplotlib, pyexcel, pyexcel-xls, pyexcel-xlsx, tkinter (tk)

#### 2. Installation
To install place the .py files into the site-package folder withing your python envrionment folder.

#### 3. Battery Data Processing
Designed for quickly processing and plotting battery data outputted by each type of battery cycler in the wagner lab.
##### 3.1. Classes
This script has three different classes available, one for each type of battery cycler in the wagner lab.
1. **arbin_file():**
This class is called when working with .xls files outputted by the arbin cycler.
2. **neware_file():**
3. **maccor_file():**

To generate a class instance to a variable using arbin_file as an example:
```Python
arbin_data = arbin_file()
```
Each instance of the class generates the following variables:
1. info: dataframe listing the channels used to run each cell, the start date, and the procedure file used.
   ```Python
   arbin.info
   ```
2. sheetnames: a list of all sheetnames in the loaded .xls file.
   ```Python
   arbin.sheetnames
   ```
3. battery_dict: a dictionary of all sheets in the .xls file.
   ```Python
   arbin.battery_dict
   ```
##### 3.2. Functions
* **load_masses**(masses, cells): calculates capacities in mAh/g, coulombic efficiency, and dQ/dV for specified cells.
  * masses is a list loading masses for the cells you want to process.
  * cells is a list of cells you want to process. Default is 'all'.
* **plot_voltage_profile**(cells, cycles, axis_range, legend_loc): plots the discharge-charge potential curves for the specified cells.
  * cells is a list of cells you want to process. Default is all.
  * cycles is a list of the cycles to plot. Default is 'all'.
  * axis range as a list [x1, x2, y1, y2]
* **plot_cycle_life**(cells, plot_type, y1_range, y2_range, decimals): plots cycle life or coulombic efficiency or both for specified cells.
  * cells is a list of cells you want to process. Default is all.
  * plot type: 'capacity', 'coulombic', or 'both'. Default is both. 
    * If 'both' selected coulombic plotted on y2 axis.
  * y1_range: left-hand axis
  * y2_range: right hand axis
  * decimals: number of decimal places axis labels are printed out to.
* **plot_dqdv**()
* **igor_datafile**()

