#Read pruned data from 'prunedData.xlsx & prunedDataOther.xlsx' and make it into the arrays that are inputs for QC_IOT.py

#imports
import pandas as pd
import os
import numpy as np

#directory
# Change working directory (Étienne MacOS)
os.chdir('/Users/Etguer/PycharmProjects/QC_IOT_data/input')  # need to use '/' or '//' instead of '\'
# Change working directory (Étienne laptop)
# os.chdir('C:/Users/user/PycharmProjects/MEDEAS_QC_data')
wd = os.getcwd()

#declare arrays & other variables
years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])
# array for already intérieur data for World (2004-2014)
asICworldArrayP = np.full(shape=(56, 56, 11), fill_value=np.nan, dtype=float)
# array for all sectors total production WORLD (2004-2014)
asTPworldArrayP = np.full(shape=(56, 1, 11), fill_value=-1, dtype=float)
# array for all sectors world value added
asVAworldArrayP = np.full(shape=(56, len(years)), fill_value=0, dtype=float)

# array for all sectors Gross Fixed Capital Formation (added to single column) WORLD
asGFCFwArrayP = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors Household demand (added to single column) WORLD
asHHwArrayP = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors Government Expenditure (added to single column) WORLD
asGEwArrayP = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors Inventories (added to single column) WORLD
asINVwArrayP = np.full(shape=(56, len(years)), fill_value=0, dtype=float)

# array for all canada sectors intermediate exports and final demand
# (CAN has 56 sectors in IE, 5 categories in final demand)
asIEFEcanArrayP = np.full(shape=(56, 2685 - 56 - 5, len(years)), fill_value=0, dtype=float)

# array for all sectors World minus QC (initially this is World and then QC is subtracted) Final Demand
asFDrowArrayP = np.full(shape=(56, len(years)), fill_value=0, dtype=float)


#read excel loop
#file name 1
str1 = 'prunedData'
str3 = '.xlsx'
#file name 2
str4 = 'Other'

row_init = 1
year_one_w = 2004

years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])
for year in years:
    yearIndex = year - year_one_w
    str2 = str(year)
    path = str1 + str2 + str3
    # number of sheets
    xl = pd.ExcelFile(path)
    sheets = len(xl.sheet_names)
    for sheetno in range(0, sheets):
        print('Reading sheet ' + xl.sheet_names[sheetno] + ' in ' + str(path)) #this should print which sheet is being read
        skiprows = row_init - 1
        tempdf = pd.read_excel(path, sheet_name=sheetno, index_col=None,
                                   header=None, skiprows=skiprows)
                                       #, usecols='', nrows=2464)
        tempar = tempdf.to_numpy().reshape(len(tempdf.index), len(tempdf.columns))
        if xl.sheet_names[sheetno] == 'asICworldArray':
            asICworldArrayP[:, :, yearIndex] = tempar
        if xl.sheet_names[sheetno] == 'asTPworldArray':
            asTPworldArrayP[:, :, yearIndex] = tempar
        if xl.sheet_names[sheetno] == 'asIEFEcanArray':
            asIEFEcanArrayP[:, :, yearIndex] = tempar

# read prunedDataOther.xlsx
path = str1 + str4 + str3
# number of sheets
xl = pd.ExcelFile(path)
sheets = len(xl.sheet_names)
for sheetno in range(0, sheets):
    print('Reading sheet ' + xl.sheet_names[sheetno] + ' in ' + str(path)) #this should print which sheet is being read
    skiprows = row_init - 1
    tempdf = pd.read_excel(path, sheet_name=sheetno, index_col=None,
                               header=None, skiprows=skiprows)
    tempar = tempdf.to_numpy().reshape(len(tempdf.index), len(tempdf.columns))
    if xl.sheet_names[sheetno] == 'asVAworldArray':
        asVAworldArrayP[:, :] = tempar
    if xl.sheet_names[sheetno] == 'asGFCFwArray':
        asGFCFwArrayP[:, :] = tempar
    if xl.sheet_names[sheetno] == 'asHHwArray':
        asHHwArrayP[:, :] = tempar
    if xl.sheet_names[sheetno] == 'asGEwArray':
        asGEwArrayP[:, :] = tempar
    if xl.sheet_names[sheetno] == 'asINVwArray':
        asINVwArrayP[:, :] = tempar
    if xl.sheet_names[sheetno] == 'asFDrowArray':
        asFDrowArrayP[:, :] = tempar
