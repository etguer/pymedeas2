# This script first reads back pruned data produced with QC_IOT_pruningWIOT.py
# This data is saved in separate arrays (ending with 'D') and then linked back to the original arrays (without the end 'D')
#
# Then the script more generally reads Québec Input-Output Tables (ES symetriques provinciaux S QC) from 2004-2018
# (PBM (Prix de Base Modifié)/Base, Intérieur, ImportationsTotales sheets)

# Takes into account that the OLD data (2004-2008)
# has different sectors than the MID data (2009-2011)
# and then the NEW data (2014-2018)

# Removes, in OLD, the three F1,F2,F3 sectors by reallocating them (according to Monica's advice)
# Reallocates and reaggregates sectors for row and col totals AND for all the data
# so that they match the ISIC classification of the WIOT used by MEDEAS-World
# calculates or extracts four matrices (A technical coefficients)
# (QC-QC Intermediate Consumption, QC-RoW Intermediate Exports,
# RoW-QC Intermediate Imports, RoW-RoW Intermediate Consumption) used in MEDEAS,
# and other economic data as in economy.xlsx

# import packages
import pandas as pd
import numpy as np
import os.path
import openpyxl
from os import listdir
from os.path import isfile, join
from openpyxl import Workbook as wb
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import quote_sheetname, absolute_coordinate, get_column_letter


def written_range(df: pd.DataFrame, startrow, startcol, header=True, index=True):
    """Return the A1 range that df.to_excel(...) writes for the given start cell.

    startrow/startcol are 0-indexed (as pandas to_excel expects).
    data_only=True returns just the data cells (excludes header rows and index cols).
    """
    h = df.columns.nlevels if header else 0
    i = df.index.nlevels if index else 0

    top_row = startrow + 1 + h 
    left_col = startcol + 1 + i 
    bot_row = startrow + h + len(df)
    right_col = startcol + i + len(df.columns)

    return f"{get_column_letter(left_col)}{top_row}:{get_column_letter(right_col)}{bot_row}"

def add_defined_name_section(writer, sheetname, name, data, startrow, startcol,header=True, index=True):
    rng = written_range(data, startrow, startcol, header=header, index=index)
    ws = writer.sheets[sheetname]
    ref = f"{quote_sheetname(ws.title)}!{absolute_coordinate(rng)}"
    ws.defined_names.add(DefinedName(name, attr_text=ref))


#=======================================================================================================================
#Read pruned data from 'prunedData.xlsx & prunedDataOther.xlsx' and make it into the arrays that are inputs for QC_IOT.py
#=======================================================================================================================
#directory
# Change working directory (Étienne MacOS)
os.chdir('/Users/Etguer/pymedeasQCdata/QC_IOT_data/QC_IOTpy_input')  # need to use '/' or '//' instead of '\'
# Change working directory (Étienne laptop)
# os.chdir('C:/Users/user/PycharmProjects/MEDEAS_QC_data')
wd = os.getcwd()

#declare arrays (pruned) & other variables
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

years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014])
                  #2015, 2016, 2017, 2018])
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
    tempdf = pd.read_excel(path, sheet_name=sheetno, index_col=None, header=None, skiprows=skiprows)
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
#=======================================================================================================================
# Change working directory (Étienne MacOS)
#os.chdir('/Users/Etguer/pymedeasQCdata/QC_IOT_data')  # need to use '/' or '//' instead of '\'
# Change working directory (Étienne laptop)
# os.chdir('C:/Users/user/PycharmProjects/MEDEAS_QC_data')
wd = os.getcwd()

# set years
all_years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018]) # this is useful for the getyearIndex function to always use 2004-2018 inclusively

yearsString = ['2004','2005','2006','2007','2008','2009','2010','2011','2012','2013','2014','2015','2016','2017','2018']

csString = [
    'Agriculture_Forestry',
    'Manufacturing',
    'Mining',
    'Utilities',
    'Construction',
    'Motor_Vehicles',
    'Wholesale_Trade',
    'Retail_Trade',
    'Hotel_Restaurants',
    'Transport',
    'Post_Telecomm',
    'Offices',
    'Education',
    'Health_SocialWork',
    'Other_services',
    'Private_Households',
]

# max number of sectors (22/25 for 2004-2008 ; 32 for 2009-2017)
maxNoSectors = 32
oldPBMNoSectors = 25

# define data structures used further
rawdata = pd.DataFrame(data=None)
rawdata_list = list()

# define yeartype ('old','mid','new') as empty variable initially
yeartype = None

# set nrows/cols for 2004-2008 & 2009-2017 for intermediate consumption (IC)
rcOldIC1 = 0
rcOldIC2 = 25
rcNewIC1 = 0
rcNewIC2 = 32

# DECLARE all the shizzle needed for this simple operation
# create each MEDEAS/WIOD2013 sector out of
commonSectorsString = ['Agriculture & Forestry', 'Manufacturing', 'Mining', 'Utilities', 'Construction',
                       'Motor Vehicles', 'Wholesale Trade', 'Retail Trade', 'Hotels and Restaurants', 'Transport',
                       'Post and Telecomm', 'Offices', 'Education', 'Health and Social work', 'Other c.s.p. services',
                       'Private Households']
QC_xi = {  # [Old, New] ; indices with decimals refer to portions of the same OLD sector
    'recycling': [15.1, 20.1],
    'minus recycling': [15.2, 20.2],
    'retail_cars': [10.1, 14.1],
    'minus retail_cars': [10.2, 14.2],
    'wholesale_cars': [9.1, 13.1],
    'minus wholesale_cars': [9.2, 13.2],
    'auto_repair': [20.1, 25.1],
    'private_households': [20.2, 25.2],
    'minus arepair&phh': [20.3, 25.3],
    'post_cour_mess': [11.1, 15.1],
    'minus post': [11.2, 15.2],
    'telecomm': [12.1, 16.1],
    'paper_publishing': [12.2, 16.2],
    'minus telecomm&paperpub': [12.3, 16.3],
    'education': [21.1, 26.1],
    'health_socialwork': [21.2, 26.2],
    'art_nonprof_rel': [21.3, 26.3]
}
# QC_IOT indices to build the 16 common sectors (indices for after deleting F1-F3, so old IOTs have 22 rows/cols)
# this is basically used as an y = a + bX2 + cX3 +dX4 function; where y = NEW sector row/col total ;
# a,b,c,d = the value of an OLD row/col total ; X1,X2,X3 the proportions to be applied to each OLD sector
# (See WorkPlan data google sheets)
# indices with decimals refer to portions of the same OLD sector
common1 = {1: [[1, 2, 3, 4], [1, 2, 3, 4]], 2: [[8], [12]], 3: [[5], [5]], 4: [[6], [6]], 5: [[7], [7, 8, 9, 10, 11]],
           6: [[], []], 7: [[], []], 8: [[], []], 9: [[19], [24]], 10: [[], []], 11: [[], []],
           12: [[13, 14, 22], [17, 18, 19, 29, 30, 31, 32]], 13: [[16], [21, 27]], 14: [[17], [22, 28]],
           15: [[18], [23]], 16: [[], []]}
common2 = {1: [[], []], 2: [[15.1], [20.1]], 3: [[], []], 4: [[], []], 5: [[], []],
           6: [[9.1], [13.1]], 7: [[9.2], [13.2]], 8: [[10.2], [14.2]], 9: [[], []], 10: [[11.2], [15.2]],
           11: [[12.1], [16.1]],
           12: [[15.2], [20.2]], 13: [[21.1], [26.1]], 14: [[21.2], [26.2]],
           15: [[21.3], [26.3]], 16: [[20.2], [25.2]]}
common3 = {1: [[], []], 2: [[12.2], [16.2]], 3: [[], []], 4: [[], []], 5: [[], []],
           6: [[10.1], [14.1]], 7: [[], []], 8: [[], []], 9: [[], []], 10: [[], []], 11: [[11.1], [15.1]],
           12: [[], []], 13: [[], []], 14: [[], []],
           15: [[12.3], [16.3]], 16: [[], []]}
common4 = {1: [[], []], 2: [[], []], 3: [[], []], 4: [[], []], 5: [[], []],
           6: [[20.1], [25.1]], 7: [[], []], 8: [[], []], 9: [[], []], 10: [[], []], 11: [[], []],
           12: [[], []], 13: [[], []], 14: [[], []],
           15: [[20.3], [25.3]], 16: [[], []]}
commonSectorsTotals = np.full(shape=(2, len(commonSectorsString), len(all_years)), fill_value=-1, dtype=float)
commonSectorsFDTotals = np.full(shape=(len(commonSectorsString), len(all_years)), fill_value=-1, dtype=float)
# QC IOT (OLD) proportions (X1,X2,X3) to commonSectors (NEW) (2004-2017; source: RAS googlesheet in Workplan Data)
pE = {
    'recycling': [2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07, 2.07],
    'minus recycling': [97.93, 97.93, 97.93, 97.93, 97.93, 97.93, 97.93, 97.93, 97.93, 97.93, 97.93, 97.93, 97.93,
                        97.93, 97.93],
    'retail_cars': [19.73, 19.57, 19.98, 17.24, 20.60, 18.91, 17.80, 18.50, 19.04, 19.31, 20.01, 20.36, 21.31, 22.04,
                    22.04],
    'minus retail_cars': [80.27, 80.43, 80.02, 82.76, 79.4, 81.09, 82.2, 81.5, 80.96, 80.69, 79.99, 79.64, 78.69,
                          77.96, 77.96],
    'wholesale_cars': [7.42, 7.55, 7.54, 6.54, 6.68, 6.43, 0.62, 6.15, 5.90, 5.70, 5.79, 6.49, 7.07, 7.04, 7.04],
    'minus wholesale_cars': [92.58, 92.45, 92.46, 93.46, 93.32, 93.57, 99.38, 93.85, 94.1, 94.3, 94.21, 93.51, 92.93,
                             92.96, 92.96],
    'auto_repair': [16.12, 16.04, 16.15, 16.35, 16.16, 16.96, 16.88, 16.88, 16.87, 17.24, 16.85, 16.22, 17.84, 17.92,
                    17.92],
    'private_households': [9.35, 8.97, 8.94, 8.93, 9.70, 10.24, 10.24, 10.29, 10.19, 10.25, 10.56, 10.53, 10.48, 10.11,
                           10.11],
    'minus arepair&phh': [74.53, 74.99, 74.91, 74.72, 74.14, 72.8, 72.88, 72.83, 72.94, 72.51, 72.59, 73.25, 71.68,
                          71.97, 71.97],
    'telecomm': [43.91, 42.94, 41.13, 40.81, 42.25, 43.23, 40.60, 40.43, 40.26, 36.82, 33.85, 32.42, 33.45, 34.70,
                 34.70],
    'post_cour_mess': [13.44, 14.21, 14.00, 14.00, 13.93, 14.83, 14.94, 14.25, 13.39, 13.32, 13.38, 12.90, 11.99,
                       12.29, 12.29],
    'minus post': [86.56, 85.79, 86.0, 86.0, 86.07, 85.17, 85.06, 85.75, 86.61, 86.68, 86.62, 87.1, 88.01, 87.71, 87.71],
    'education': [33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33],
    'health_socialwork': [33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33],
    'art_nonprof_rel': [33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33, 33.33],
    'paper_publishing': [15.38, 16.04, 16.47, 17.77, 16.62, 15.81, 15.15, 13.49, 12.96, 13.89, 14.18, 14.05, 12.48,
                         11.96, 11.96],
    'minus telecomm&paperpub': [40.71, 41.02, 42.4, 41.42, 41.13, 40.96, 44.25, 46.08, 46.78, 49.29, 51.97, 53.53,
                                54.07, 53.34, 53.34],
    'None': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
}


# set either old, mid or new as TRUE to be used to define which year period parameters to use
def setyeartype(yr):  # yr = year
    old_local = 2004 <= yr <= 2008  # if year is between 2004-2008, OLD datasets with F1, F2, F3 to be reallocated
    mid_local = 2009 <= yr <= 2011
    new_local = 2012 <= yr <= 2018
    global yeartype
    if old_local:
        yeartype = 'old'
    else:
        if mid_local:
            yeartype = 'mid'
        else:
            if new_local:
                yeartype = 'new'


# get yearIndex (2004 = 0, 2018 = 13)
def getyearindex(yrs, yr):
    output = np.where(yrs == yr)[0][0]  # index of year in years
    return output


# import data function
# raw data for all years (all rows and all columns with numbers from ES symetriques provinciaux S QC)
def getrawdata(dr, yr, data_name):  # dr = directory; yr = year; data_name = PBM/intérieur if empty, otherwise can be
    # ImportationsTotales, ...
    stringname0 = 'WIOT_ROW/'
    stringname1 = 'TableauxSymetriquesQC/ES symetriques provinciaux S QC '
    stringname3 = '.xls'
    stringname4 = '.xlsx'
    stringname1_1 = 'WIOT'
    stringname1_3 = '_Nov16_ROW'
    stringname1_4 = '.xlsb'
    row_skip = ''
    all_cols = ""
    stringname = ""
    sheet_name = ""
    nrows = ""
    if data_name == 'TP_World':
        print('getrawdata:TP_World')
        stringname = stringname0 + stringname1_1 + str(yr) + stringname1_3 + stringname1_4
        sheet_name = str(yr)
        all_cols = 'CYK'
        row_skip = 6
        nrows = None
    else:
        if data_name == 'Canada_Exports':
            print('getrawdata:Canada_Exports')
            stringname = stringname0 + stringname1_1 + str(yr) + stringname1_3 + stringname1_4
            sheet_name = str(yr)
            all_cols = 'E:CYK'  # intermediate export + final export
            row_skip = 286  # for CANADA only
            nrows = 56  # for CANADA only
        else:
            if data_name == 'ImportationsTotales':
                print('getrawdata:ImportationsTotales')
                if yeartype == 'old':
                    pass
                else:
                    if yeartype == 'mid':
                        row_skip = range(0, 3)
                        stringname = stringname1 + str(yr) + stringname3
                        all_cols = 'D:BC'
                        sheet_name = 'ImportationsTotales'
                        nrows = None
                    else:
                        if yeartype == 'new':
                            stringname = stringname1 + str(yr) + stringname4
                            row_skip = range(0, 3)
                            all_cols = 'D:BE'
                            sheet_name = 'ImportationsTotales'
                            nrows = None
            else:
                if data_name == '':
                    print('getrawdata:''')
                    if yeartype == 'old':
                        stringname = stringname1 + str(yr) + stringname3
                        row_skip = range(0, 3)
                        sheet_name = 'PBM'
                        all_cols = 'D:BP'  # columns to read from raw data sets
                        nrows = None
                    else:
                        if yeartype == 'mid':
                            stringname = stringname1 + str(yr) + stringname3
                            row_skip = range(0, 3)
                            sheet_name = 'Intérieur'
                            all_cols = 'D:CC'  # columns to read from raw data sets
                            nrows = None
                        else:
                            if yeartype == 'new':
                                stringname = stringname1 + str(yr) + stringname4
                                row_skip = range(0, 3)
                                sheet_name = 'Intérieur'
                                all_cols = 'D:CG'  # columns to read from raw data sets
                                nrows = None
                else:
                    if data_name == 'Base':
                        print('getrawdata:Base')
                        if yeartype == 'old':
                            stringname = stringname1 + str(yr) + stringname3
                            row_skip = range(0, 3)
                            sheet_name = 'PBM'
                            all_cols = 'D:BP'  # columns to read from raw data sets
                            nrows = None
                        else:
                            if yeartype == 'mid':
                                stringname = stringname1 + str(yr) + stringname3
                                row_skip = range(0, 3)
                                sheet_name = 'Base'
                                all_cols = 'D:CC'  # columns to read from raw data sets
                                nrows = None
                            else:
                                if yeartype == 'new':
                                    stringname = stringname1 + str(yr) + stringname4
                                    row_skip = range(0, 3)
                                    sheet_name = 'Base'
                                    all_cols = 'D:CG'  # columns to read from raw data sets
                                    nrows = None
                    else:
                        if data_name == 'Quebec_Exports':
                            print('getrawdata:Quebec_Exports')
                            if yeartype == 'old':
                                stringname = stringname1 + str(yr) + stringname3
                                sheet_name = 'PBM'
                                row_skip = range(0, 3)
                                all_cols = 'AM:AN,AP:BB'
                                nrows = None
                            else:
                                if yeartype == 'mid':
                                    stringname = stringname1 + str(yr) + stringname3
                                    sheet_name = 'Base'
                                    row_skip = range(0, 3)
                                    all_cols = 'AZ:BA,BC:BO'
                                    nrows = None
                                else:
                                    if yeartype == 'new':
                                        stringname = stringname1 + str(yr) + stringname4
                                        sheet_name = 'Base'
                                        row_skip = range(0, 3)
                                        all_cols = 'BB:BC,BE:BR'
                                        nrows = None
                        else:
                            if data_name == 'FD_World':
                                print('getrawdata:FD_World')
                                stringname = stringname0 + stringname1_1 + str(yr) + stringname1_3 + stringname1_4
                                sheet_name = str(yr)
                                all_cols = 'CPY:CYJ'
                                row_skip = 6
                                nrows = None
                            else:
                                print('data_name', data_name, 'is not a valid argument')
                                return None
    if os.path.isfile(dr + '/' + stringname):  # check if file exists in working directory
        print('Reading', sheet_name, 'in', stringname, 'in', dr)
        # read file, remove index and column names, only keep numbers (0-XX)
        print('stringname:',stringname,'sheet_name:',sheet_name,'skiprows:',row_skip,'usecols:',all_cols,'nrows:',nrows)
        tempdata = pd.read_excel(stringname, sheet_name=sheet_name, index_col=None,
                                 header=None, skiprows=row_skip, usecols=all_cols, nrows=nrows)
        tempdata = tempdata.reset_index(drop=True)  # reset row index
        tempdata.columns = range(tempdata.shape[1])  # reset column index
        return tempdata
    else:
        print(data_name, 'does not exist in', stringname)
        return


# make dataframe section into multi-year array (data argument can be IC, --- FD, TP, TI)
def df_to_array(df_list, data_name, yr):
    row1 = None
    row2 = None
    subset = None
    data_name_str = data_name
    yeari = getyearindex(all_years, yr)
    df = df_list[yeari]
    # make dataframe into an array
    if df is not None:
        darr = df.to_numpy().reshape(len(df.index), len(df.columns))
        darr = np.where(darr == '.', 0, darr)
        if data_name == 'TP_World':
            subset = darr[:]
            temp = np.full(shape=56, fill_value=0, dtype=float)
            # aggregate across 43 countries, each spaced by 56 sectors
            for j in range(0, 43+1):
                temp[:] = temp[:] + subset[(j * 56):(55+1 + (56 * j)), 0]  # needs 2009
            global asTPwArray
            asTPwArray[:, yeari] = temp[:]
        if data_name == 'IE_CAN':
            exclude1 = range(2489, 2493 + 1)
            exclude2 = range(280, 335 + 1)
            subset = np.delete(darr, exclude1, axis=1)
            subset = np.delete(subset, exclude2, axis=1)
            global asIEFEcanArray
            asIEFEcanArray[:, :, yeari] = subset[:, :]
        if data_name == 'TE':
            if yeartype == 'old':
                row1 = rcOldIC1  # all sectors rows for exports
                row2 = rcOldIC2
            else:
                if yeartype == 'mid' or 'new':
                    row1 = rcNewIC1
                    row2 = rcNewIC2
            subset = np.sum(darr[row1:row2, :], axis=1)
            global asTEArray
            asTEArray[row1:row2, yeari] = subset[:]
        if data_name == 'CC':
            if yeartype == 'old':
                rows = [35, 36] - (1+3)*np.ones(2, dtype=int)  # -1 to start at 0, -3 for skiprows
                col1 = rcOldIC1
                col2 = rcOldIC2
            else:
                if yeartype == 'mid':
                    rows = [56, 57] - (1+3)*np.ones(2, dtype=int)  # -1 to start at 0, -3 for skiprows
                    col1 = rcNewIC1
                    col2 = rcNewIC2
                else:
                    if yeartype == 'new':
                        rows = [57, 58] - (1 + 3) * np.ones(2, dtype=int)  # -1 to start at 0, -3 for skiprows
                        col1 = rcNewIC1
                        col2 = rcNewIC2
            subset = np.sum(darr[rows, col1:col2], axis=0)
            global asCCArray
            asCCArray[0:len(subset), yeari] = subset[:]
        if data_name == 'VA':
            if yeartype == 'old':
                rows = [29, 30, 31, 32, 33, 34, 35, 36] - (1+3)*np.ones(8, dtype=int)  # -1 to start at 0, -3 for skiprows
                col1 = rcOldIC1
                col2 = rcOldIC2
            else:
                if yeartype == 'mid':
                    rows = [50, 51, 52, 53, 54, 55, 56, 57] - (1 + 3) * np.ones(8, dtype=int)  # -1 to start at 0, -3 for skiprows
                    col1 = rcNewIC1
                    col2 = rcNewIC2
                else:
                    if yeartype == 'new':
                        rows = [51, 52, 53, 54, 55, 56, 57, 58] - (1 + 3) * np.ones(8, dtype=int)  # -1 to start at 0, -3 for skiprows
                        col1 = rcNewIC1
                        col2 = rcNewIC2
            subset = np.sum(darr[rows, col1:col2], axis=0)
            global asVAArray
            asVAArray[0:len(subset), yeari] = subset[:]
        if data_name == 'LC':
            if yeartype == 'old':
                rows = [33, 34] - (1+3)*np.ones(2, dtype=int)  # -1 to start at 0, -3 for skiprows
                col1 = rcOldIC1
                col2 = rcOldIC2
            else:
                if yeartype == 'mid':
                    rows = [54, 55] - (1+3)*np.ones(2, dtype=int)  # -1 to start at 0, -3 for skiprows
                    col1 = rcNewIC1
                    col2 = rcNewIC2
                else:
                    if yeartype == 'new':
                        rows = [55, 56] - (1 + 3) * np.ones(2, dtype=int)  # -1 to start at 0, -3 for skiprows
                        col1 = rcNewIC1
                        col2 = rcNewIC2
            subset = np.sum(darr[rows, col1:col2], axis=0)
            global asLCArray
            asLCArray[0:len(subset), yeari] = subset[:]
        if data_name == 'GFCF':
            if yeartype == 'old':
                cols = [27, 28, 29, 30, 31, 32] - (1)*np.ones(6, dtype=int)  # -1 to start at 0
                row1 = rcOldIC1
                row2 = rcOldIC2
            else:
                if yeartype == 'mid':
                    cols = [36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46] - (1)*np.ones(11, dtype=int)  # -1 to start at 0
                    row1 = rcNewIC1
                    row2 = rcNewIC2
                else:
                    if yeartype == 'new':
                        cols = [40, 41, 42, 43, 44, 45, 46, 47, 48, 49] - (1) * np.ones(10, dtype=int)  # -1 to start at 0
                        row1 = rcNewIC1
                        row2 = rcNewIC2
            subset = np.sum(darr[row1:row2, cols], axis=1)
            global asGFCFArray
            asGFCFArray[0:len(subset), yeari] = subset[:]
        if data_name == 'HH':
            if yeartype == 'old':
                cols = [26] - (1)*np.ones(1, dtype=int)  # -1 to start at 0
                row1 = rcOldIC1
                row2 = rcOldIC2
            else:
                if yeartype == 'mid':
                    cols = [33] - (1)*np.ones(1, dtype=int)  # -1 to start at 0
                    row1 = rcNewIC1
                    row2 = rcNewIC2
                else:
                    if yeartype == 'new':
                        cols = [33, 34, 35, 36, 37] - (1) * np.ones(5, dtype=int)  # -1 to start at 0
                        row1 = rcNewIC1
                        row2 = rcNewIC2
            subset = np.sum(darr[row1:row2, cols], axis=1)
            global asHHArray
            asHHArray[0:len(subset), yeari] = subset[:]
        if data_name == 'GE':
            if yeartype == 'old':
                cols = [35] - (1)*np.ones(1, dtype=int)  # -1 to start at 0
                row1 = rcOldIC1
                row2 = rcOldIC2
            else:
                if yeartype == 'mid':
                    cols = [34, 35] - (1)*np.ones(2, dtype=int)  # -1 to start at 0
                    row1 = rcNewIC1
                    row2 = rcNewIC2
                else:
                    if yeartype == 'new':
                        cols = [38, 39] - (1) * np.ones(2, dtype=int)  # -1 to start at 0
                        row1 = rcNewIC1
                        row2 = rcNewIC2
            subset = np.sum(darr[row1:row2, cols], axis=1)
            global asGEArray
            asGEArray[0:len(subset), yeari] = subset[:]
        if data_name == 'INV':
            if yeartype == 'old':
                cols = [33, 34] - (1)*np.ones(2, dtype=int)  # -1 to start at 0
                row1 = rcOldIC1
                row2 = rcOldIC2
            else:
                if yeartype == 'mid':
                    cols = [47, 48] - (1)*np.ones(2, dtype=int)  # -1 to start at 0
                    row1 = rcNewIC1
                    row2 = rcNewIC2
                else:
                    if yeartype == 'new':
                        cols = [50] - (1) * np.ones(1, dtype=int)  # -1 to start at 0
                        row1 = rcNewIC1
                        row2 = rcNewIC2
            subset = np.sum(darr[row1:row2, cols], axis=1)
            global asINVArray
            asINVArray[0:len(subset), yeari] = subset[:]
        if data_name == 'FDw':
            subset = np.sum(darr[:, :], axis=1)
            temp = np.full(shape=56, fill_value=0, dtype=float)
            # aggregate across 43 countries, each spaced by 56 sectors
            for j in range(0, 43+1):
                temp[:] = temp[:] + subset[(j * 56):(55+1 + (56 * j))]  # needs 2009
            global asFDrowArray
            asFDrowArray[:, yeari] = temp[:]
        if data_name == 'IC':  # topleft Intermediate Consumption matrix
            if yeartype == 'old':
                # set row and column range to subset according to yeartype
                row1 = rcOldIC1
                row2 = rcOldIC2
                col1 = rcOldIC1
                col2 = rcOldIC2
                # subset IC from array
                subset = darr[row1:row2, col1:col2]
                # update global array for that year
                global asIC_PBM_Array
                asIC_PBM_Array[row1:row2, col1:col2, yeari] = subset[:, :]
            else:
                if yeartype == 'mid' or 'new':
                    # set row and column range to subset according to yeartype
                    row1 = rcNewIC1
                    row2 = rcNewIC2
                    col1 = rcNewIC1
                    col2 = rcNewIC2
                    # subset IC from array
                    subset = darr[row1:row2, col1:col2]
                    # update global array for that year
                    global asICArray
                    asICArray[row1:row2, col1:col2, yeari] = subset[:, :]

        if data_name == 'II':  # bottomleft Intermediate Imports matrix
            if yeartype == 'old':  # these will be estimated using the proportions from the mid/new data
                pass
            else:
                if yeartype == 'mid' or 'new':
                    # set row and column range to subset according to yeartype
                    row1 = rcNewIC1  # same for II and IC
                    row2 = rcNewIC2
                    col1 = rcNewIC1
                    col2 = rcNewIC2
                    # subset IC from array
                    subset = darr[row1:row2, col1:col2]
                    # update global array for that year
                    global asIIArray
                    asIIArray[row1:row2, col1:col2, yeari] = subset[:, :]
        if data_name == 'TP':  # top right Total Production vector
            if yeartype == 'old':
                # set row and column range to subset according to yeartype
                col1 = 64
                row1 = rcOldIC1
                row2 = rcOldIC2
                # subset IC from array
                subset = darr[row1:row2, col1]
            if yeartype == 'mid':
                # set row and column range to subset according to yeartype
                col1 = 77
                row1 = rcNewIC1
                row2 = rcNewIC2
                # subset IC from array
                subset = darr[row1:row2, col1]
            if yeartype == 'new':
                # set row and column range to subset according to yeartype
                col1 = 81
                row1 = rcNewIC1
                row2 = rcNewIC2
                # subset IC from array
                subset = darr[row1:row2, col1]
            # update global array for that year
            global asTPArray
            asTPArray[row1:row2, yeari] = subset
    else:
        print('df_to_array', 'data_name:', data_name_str, 'data does not exist for year', yr)


# reaggregate F1, F2 & F3 (for 2004-2008 )
def f_reagg(array):
    output = None
    print('REALLOCATING F1, F2, F3')
    f = [21, 22, 23] - np.ones(3, dtype=int)
    if array.shape[0] == array.shape[1] and array.ndim == 3:  # for symetrical sector-sector arrays
        print('array is symmetrical')
        for yr in years:
            setyeartype(yr)
            if yeartype == 'old':
                yeari = getyearindex(years, yr)
                for sector in [15, 16, 20]-np.ones(3, dtype=int):
                    # 21(F1) --> 15(services prof.), 16(services admin.) & 20(autres services)
                    # 22(F2) --> 15(services prof.), 16(services admin.) & 20(autres services)
                    array[sector, :, yeari] = array[sector, :, yeari] + \
                                              array[21-1, :, yeari] / 3 + array[22-1, :, yeari] / 3
                    array[:, sector, yeari] = array[:, sector, yeari] + \
                                              array[:, 21-1, yeari] / 3 + array[:, 22-1, yeari] / 3
                # 23(F3) --> 11(transport et entreposage)
                array[11-1, :, yeari] = array[11-1, :, yeari] + array[23-1, :, yeari]
                array[:, 11-1, yeari] = array[:, 11-1, yeari] + array[:, 23-1, yeari]
        # remove F1, F2, F3 rows and columns
        array = np.delete(np.delete(array, f, axis=0), f, axis=1)  # delete F for rows, then columns
        print('array shape:', array.shape)
        output = array
    else:
        if array.ndim == 2:  # for final demand arrays (must be already summed up) or total production arrays
            print('array is not symmetrical')
            for yr in years:
                setyeartype(yr)
                if yeartype == 'old':
                    yeari = getyearindex(years, yr)
                    for sector in [15, 16, 20]-np.ones(3, dtype=int):
                        array[sector, yeari] = array[sector, yeari] + array[21-1, yeari] / 3 + \
                                               array[22-1, yeari] / 3
                    # 23(F3) --> 11(transport et entreposage)
                    array[11-1, yeari] = array[11-1, yeari] + array[23-1, yeari]
            # remove F1, F2, F3 rows ***** for cols 0:4 (2004-2008) only *****
            temparray = np.full(shape=(array.shape[0], 5), fill_value=-1, dtype=float)
            temparray[:, 0:5] = array[:, 0:5]
            temparray = np.delete(temparray, f, axis=0)
            array[:, 0:5] = 0
            array[0:temparray.shape[0], 0:5] = temparray[:, 0:5]
            output = array
    return output


# function to return key for any value
def gk(val, oldnewgk):
    for key, value in QC_xi.items():
        if len(val[oldnewgk]) > 0:
            if val[oldnewgk][0] == value[oldnewgk]:
                if key is None:
                    key = 'None'
                return key

# get sum when rows and cols arrays have many values (common1)
def getsums(array, rows, cols, yr):
    output = 0
    for r in rows:
        for c in cols:
            output = output + array[r - 1, c - 1, yr]
    return output


# asICArray -> csICArray Currently works for asICArray (22*22 (2004-2008) or 32*32 (2009-2018))
def as_to_cs(array):
    finalarray = None
    # ----- QC_IOT sectoral totals => Common sectoral totals -----
    #         commonSectorsTotals[0, :, :] =≥ rows
    #         commonSectorsTotals[1, :, :] =≥ cols
    print('BEGIN  as (all sectors) (22)===>>> WIOT cs (common sectors) (16) REALLOCATION')
    if array.ndim == 3:
        finalarray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), len(years)), fill_value=-1,
                             dtype=float)
    else:
        if array.ndim == 2:
            finalarray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=-1, dtype=float)

    for year in years:
        yeari = getyearindex(years, year)
        setyeartype(year)
        oldnew = None
        if yeartype == 'old':
            oldnew = 0
        else:
            if yeartype == 'mid' or 'new':
                oldnew = 1
        print('*** START FILLING cs array year/yeari', year, yeari, '***')
        if array.ndim == 3:
            for row in range(0, 16):
                #print('========================================ROW:', row)
                for col in range(0, 16):
                    #print('========================================COL:', col)
                    tempAggC = np.zeros(shape=(4, 4), dtype=float)
                    for cmnRow in range(0, 4):
                        #print('cmnR:', cmnRow+1)
                        commonRow = [common1, common2, common3, common4][cmnRow]
                        tempcRow = commonRow[row + 1][oldnew]
                        #print('commonRow[row+1][oldnew]:', commonRow[row+1][oldnew])
                        for cmnCol in range(0, 4):
                            #print('cmnC:', cmnCol+1)
                            commonCol = [common1, common2, common3, common4][cmnCol]
                            tempcCol = commonCol[col + 1][oldnew]

                            if len(tempcRow) == 0:
                                tempAggC[cmnRow, cmnCol] = 0
                                #print('tempAggC', cmnRow+1, cmnCol+1, ': ZERO!!!')
                            else:
                                if len(tempcCol) == 0:
                                    tempAggC[cmnRow, cmnCol] = 0
                                    #print('tempAggC', cmnRow+1, cmnCol+1, ': ZERO!!!')
                                else:
                                    if cmnRow + 1 == 1 and cmnCol + 1 == 1:
                                        #print('tempcRow:', tempcRow)
                                        #print('tempcCol:', tempcCol)
                                        tempAggC[cmnRow, cmnCol] = getsums(array,
                                                                           tempcRow,
                                                                           # - np.ones(len(tempcRow), dtype=int)
                                                                           tempcCol,
                                                                           # - np.ones(len(tempcCol), dtype=int)
                                                                           yeari)
                                    #                                    print('value=', tempAggC[cmnRow, cmnCol])
                                    else:
                                        if cmnRow + 1 == 1:
                                            #print('allooo')
                                            #print('tempcRow:', tempcRow)
                                            #print('tempcCol:', tempcCol)
                                            tempAggC[cmnRow, cmnCol] = \
                                                np.sum(array[tempcRow -
                                                             np.ones(len(tempcRow), dtype=int),
                                                             int(tempcCol[0]) - 1,
                                                             yeari]) \
                                                * pE[str(gk(commonCol[col + 1], oldnew))][yeari] / 100
                                            #print('pE=', pE[str(gk(commonCol[col+1], oldnew))][yeari])
                                            #print('value=', tempAggC[cmnRow, cmnCol])
                                        else:
                                            if cmnCol + 1 == 1:
                                                #print('tempcRow', tempcRow)
                                                #print('tempcCol:', tempcCol)
                                                tempAggC[cmnRow, cmnCol] = \
                                                    np.sum(array[int(tempcRow[0]) - 1,
                                                                 tempcCol - np.ones(len(tempcCol), dtype=int),
                                                                 yeari]) \
                                                    * pE[str(gk(commonRow[row + 1], oldnew))][yeari] / 100
                                                #print('pE=', pE[str(gk(commonRow[row+1], oldnew))][yeari])
                                                #print('value=', tempAggC[cmnRow, cmnCol])
                                            else:
                                                #print('tempcRow', tempcRow)
                                                #print('tempcCol:', tempcCol)
                                                tempAggC[cmnRow, cmnCol] = \
                                                    array[int(tempcRow[0]) - 1, int(tempcCol[0]) - 1, yeari] * \
                                                    pE[str(gk(commonRow[row + 1], oldnew))][yeari] / 100 * \
                                                    pE[str(gk(commonCol[col + 1], oldnew))][yeari] / 100
                                                #print('pE=', pE[str(gk(commonRow[row+1], oldnew))][yeari] *
                                                      #pE[str(gk(commonCol[col+1], oldnew))][yeari])
                                                #print('value=', tempAggC[cmnRow, cmnCol])
                                                #print('tempAggC:', tempAggC)
                    finalarray[row, col, yeari] = np.sum(tempAggC)
        else:
            if array.ndim == 2:
                #print('year:', year, 'oldnew:', oldnew)
                for cs in range(0, 16):
                    #print('cs:', cs)
                    temptotc = [0, 0, 0,
                                0]  # storage for the partial values of the common sector according to the previous classification
                    for cmn in [1, 2, 3, 4]:
                        #print('cmn:', cmn)
                        common = [common1, common2, common3, common4][cmn - 1]
                        tempc = common[cs + 1][oldnew]
                        #print('tempc:', tempc)
                        if len(tempc) == 0:
                            temptotc[cmn - 1] = 0
                            #print('temptotc', cmn, ': ZERO!!!')
                        else:
                            if cmn - 1 == 0:
                                temptotc[cmn - 1] = np.sum(array[tempc - np.ones(len(tempc), dtype=int), yeari])
                            else:
                                temptotc[cmn - 1] = array[int(tempc[0]) - 1, yeari] \
                                                    * pE[str(gk(common[cs + 1], oldnew))][yeari] / 100
                    #print('temptotc', cmn, ':', temptotc[cmn - 1])
                    finalarray[cs, yeari] = np.sum(temptotc)
    return finalarray


# check totals pre (as) and post (cs) aggregation to 16 common sectors
def checksums_as_cs(asarray, csarray):
    error = False
    print('***  CHECKING TOTALS FOR asArray -> csArray  ***')
    if csarray.ndim == 3 and asarray.ndim == 3:  # for 3d arrays
        for year in years:
            yearIndex = getyearindex(years, year)
            #            print('year', year)
            rTpostArr = np.sum(np.sum(csarray[:, :, yearIndex], axis=0), axis=0)
            rTpreArr = np.sum(np.sum(asarray[:, :, yearIndex], axis=0), axis=0)
            #            cTpostArr = np.sum(np.sum(csarray[:, :, yearIndex], axis=1), axis=0)
            #            cTpreArr = np.sum(np.sum(asarray[:, :, yearIndex], axis=1), axis=0)
            #            print('Total from rows pre allocation (Arrays):', rTpreArr)
            #            print('Total from rows post allocation (Arrays):', rTpostArr)
            #            print('Total from cols pre allocation (Arrays):', cTpreArr)
            #            print('Total from cols post allocation (Arrays):', cTpostArr)

            maxrT = max(rTpreArr, rTpostArr)
            minrT = min(rTpreArr, rTpostArr)
            #            print('maxrT:',maxrT, 'minrT:',minrT)
            pError = (maxrT - minrT) / (minrT * 100 + 1e-100)
            #            print('Row error is:', pError, '%')
            if pError > 1:
                print('Row Totals not equal to Column Totals', 'maxrT:',
                      maxrT, 'minrT:', minrT, 'Row error is:', pError, '%')
                error = True
    else:
        if csarray.ndim == 2 and asarray.ndim == 2:  # for 2d arrays such as FD, TP
            for year in years:
                yearIndex = getyearindex(years, year)
                rTpostArr = np.sum(csarray[:, yearIndex], axis=0)
                rTpreArr = np.sum(asarray[:, yearIndex], axis=0)
                maxrT = max(rTpreArr, rTpostArr)
                minrT = min(rTpreArr, rTpostArr)
                pError = (maxrT - minrT) / (minrT * 100 + 1e-100)
                if pError > 1:
                    print('Row Totals not equal to Column Totals', 'maxrT:',
                          maxrT, 'minrT:', minrT, 'Row error is:', pError, '%')
                    error = True
        else:
            if csarray.ndim != asarray.ndim:
                print('ERROR: both arrays must have same dimension')
                return
    if not error:
        print('======================== as -> cs reaggregation SUCCESSFUL =========================')


# get row and col totals per year post F1-F2-F3 reaggregation (2004-2008) and check if totals are the same as preagg
def freagg_check(yrs, array):
    # (after F1-F3 reallocation)
    rtots1 = np.zeros(shape=(maxNoSectors, len(yrs)))  # initialize arrays that store row and col totals
    ctots1 = np.zeros(shape=(maxNoSectors, len(yrs)))  # row or col total on rows, year for columns
    # difference arrays
    d_r = np.zeros(shape=len(yrs))  # initialize row difference array
    d_c = np.zeros(shape=len(yrs))  # initialize col difference array
    for yr in yrs:
        yearIndex = getyearindex(yrs, yr)
        setyeartype(yr)
        if yeartype == 'old':
            if array.ndim == 3:
                #        print('row/colTOT year',year,'yearIndex',yearIndex)
                # check row and col totals
                for row in range(0, array.shape[0]):
                    #            print('row:',row)
                    rtots1[row, yearIndex] = np.sum(array[row, :, yearIndex])
                #            print('rowtotal:', np.sum(asIC_PBM_Array[row, :, yearIndex]))
                for col in range(0, array.shape[1]):
                    #            print('col:',col)
                    ctots1[col, yearIndex] = np.sum(array[:, col, yearIndex])
                d_r[yearIndex] = np.sum(rowTotals[:, yearIndex]) - np.sum(rtots1[:, yearIndex])
                d_c[yearIndex] = np.sum(colTotals[:, yearIndex]) - np.sum(ctots1[:, yearIndex])

                # check if pre and post F aggregation totals are the same
                if (abs(d_r[yearIndex]) <= 1) and (abs(d_c[yearIndex])) <= 1:
                    print('year', yr, 'Row totals and Column totals unchanged, reallocation successful',
                          'row/col differences are:', d_r[yearIndex], '/', d_c[yearIndex])
                else:
                    print('year', yr, '*** Row totals and Column totals changed, reallocation unsuccessful',
                          'row/col differences are:', d_r[yearIndex], '/', d_c[yearIndex])
            else:
                if array.ndim == 2:
                    for row in range(0, maxNoSectors): # 'maxNoSectors' previously array.shape[0]
                        rtots1[row, yearIndex] = array[row, yearIndex]
                    d_r[yearIndex] = np.sum(rowTotals[:, yearIndex]) - np.sum(rtots1[:, yearIndex])
                    # check if pre and post F aggregation totals are the same
                    if (abs(d_r[yearIndex]) <= 1):
                        print('year', yr, 'Row total unchanged, reallocation successful',
                              'row difference are:', d_r[yearIndex])
                    else:
                        print('year', yr, '*** Row total changed, reallocation unsuccessful',
                              'row difference are:', d_r[yearIndex])

        global rcTot1
        rcTot1 = np.full(shape=(maxNoSectors, 2, len(years)), fill_value=-1, dtype=float)
        rcTot1[:, 0, yearIndex] = rtots1[:, yearIndex]  # !!! losing the decimals here, weirdly !!!
        rcTot1[:, 1, yearIndex] = ctots1[:, yearIndex]  # !!! same !!! (only in the printing I am pretty sure)


# get row and col totals per year pre F1-F2-F3 reaggregation (2004-2008)
def get_pre_rc_tots(yrs, array):
    rtots = np.zeros(shape=(maxNoSectors, len(years)))  # initialize arrays that store row and col totals
    ctots = np.zeros(shape=(maxNoSectors, len(years)))  # row or col total on rows, year for columns
    nrows = np.shape(array)[0]
    ncols = np.shape(array)[1]
    for yr in yrs:
        yearIndex = getyearindex(yrs, yr)
        setyeartype(yr)
        if yeartype == 'old':
            if array.ndim == 3:
                for row in range(0, nrows): # 'maxNoSectors' previously array.shape[0]
                    rtots[row, yearIndex] = np.sum(array[row, :, yearIndex])
                for col in range(0, ncols): # 'maxNoSectors' previously array.shape[1]
                    ctots[col, yearIndex] = np.sum(array[:, col, yearIndex])
            else:
                if array.ndim == 2:
                    for row in range(0, nrows): # 'maxNoSectors' previously array.shape[0]
                        rtots[row, yearIndex] = np.sum(array[row, yearIndex])
    global rowTotals
    rowTotals = rtots
    global colTotals
    colTotals = ctots


# calculate sd for 1 dimensional array
def get_sd(array):
    if array.ndim != 1:
        print('ERROR: array must have 1 dimension')
        return
    else:
        temp = 0
        mu = sum(array) / len(array)
        for val in array:
            temp = (val - mu) ** 2 + temp
        sd = (temp / len(array)) ** (1 / 2)
        return sd


# check that ICtot -> IC + II reagg (2004-2008) works
def checksums_ICtot_IC_II_reagg(ICtotarray, ICarray, IIarray):
    maxdiff = np.max(ICtotarray[:, :, 0:5] - ICarray[:, :, 0:5] - IIarray[:, :, 0:5])
    mindiff = np.min(ICtotarray[:, :, 0:5] - ICarray[:, :, 0:5] - IIarray[:, :, 0:5])
    if abs(maxdiff) > 1 or abs(mindiff) > 1:
        print('Error: ICtot -> IC + II reaggregation unsuccessful; largest differences are:', maxdiff, mindiff)
    else:
        print('========================= ICtot -> IC + II reaggregation successful =========================')


# get A matrices out of the Input Output 3D array (IC QC, II, IE & IC RoW) and the total production 2D array
def get_A(io_array, tp_array):
    amatrix = np.full(shape=(io_array.shape[0], io_array.shape[1], len(years)), fill_value=-1, dtype=float)
    for year in years:
        yearIndex = getyearindex(years, year)
        for col in range(0, io_array.shape[0]):
            amatrix[:, col, yearIndex] = io_array[:, col, yearIndex] / (tp_array[col, yearIndex] + 1e-100)
    return amatrix


# get L matrices out of the Input Output 3D array (IC QC, II, IE & IC RoW) and the total production 2D array
def get_L(amatrix):
    lmatrix = np.full(shape=(amatrix.shape[0], amatrix.shape[1], len(years)), fill_value=-1, dtype=float)
    imatrix = np.identity(n=amatrix.shape[0])
    for year in years:
        yearIndex = getyearindex(years, year)
        lmatrix[:, :, yearIndex] = np.linalg.inv(imatrix - amatrix[:, :, yearIndex])
        return lmatrix


# get ALL raw data
rawdataIC_list = list()  # PBM sheet for 2004-2008, Intérieur sheet for 2009-2018
# final demand FD (QC - QC final goods and services sold to consumers)
rawdataE_Canada_list = list()  # final and intermediate exports (Canada - RoW, used for proportions for Quebec)
rawdataII_list = list()  # intermediate imports (RoW to QC industry exchanges)
rawdataTE_list = list()  # intermediate and final exports (QC to RoW)
rawdataTP_World_list = list()  # total production world (for A matrix for QC IE)
rawdataBase_list = list()  # PBM/Base IO sheet for all years
rawdataFD_World_list = list() # Final demand (HH+GFCF+GE+INV) for World


years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])
for year in years:
    setyeartype(year)
    print('yeartype:', yeartype)
    print('yearIndex:', getyearindex(years, year))

    # if year == 2009:
    #     rawdataFD_World = pd.DataFrame(data=None, index=range(2472), columns=range(220))
    #     rawdataFD_World_list.append(rawdataFD_World)
    #     rawdataTP_World = pd.DataFrame(data=None, index=range(2472), columns=range(1))
    #     rawdataTP_World_list.append(rawdataTP_World)
    #     rawdataE_Canada = pd.DataFrame(data=None, index=range(56), columns=range(2685))
    #     rawdataE_Canada_list.append(rawdataE_Canada)
    # else:
    #     if year != 2009:
    #        rawdataFD_World = getrawdata(wd, year, data_name='FD_World')
    #        rawdataFD_World_list.append(rawdataFD_World)
    #        rawdataTP_World = getrawdata(wd, year, data_name='TP_World')
    #        rawdataTP_World_list.append(rawdataTP_World)
    #        rawdataE_Canada = getrawdata(wd, year, data_name='Canada_Exports')
    #        rawdataE_Canada_list.append(rawdataE_Canada)

    if year == 2012 or year == 2013:
        rawdataIC = pd.DataFrame(data=None, index=range(56), columns=range(82))
        rawdataIC_list.append(rawdataIC)
        rawdataBase = pd.DataFrame(data=None, index=range(41), columns=range(82))
        rawdataBase_list.append(rawdataBase)
        rawdataII = pd.DataFrame(data=None, index=range(41), columns=range(54))
        rawdataII_list.append(rawdataII)
        rawdataTE = pd.DataFrame(data=None, index=range(55), columns=range(15))
        rawdataTE_list.append(rawdataTE)
    else:
        if year != 2012 or year != 2013:
            rawdataIC = getrawdata(wd, year, data_name='')
            rawdataIC_list.append(rawdataIC)
            rawdataBase = getrawdata(wd, year, data_name='Base')
            rawdataBase_list.append(rawdataBase)
            rawdataII = getrawdata(wd, year, data_name='ImportationsTotales')
            rawdataII_list.append(rawdataII)
            rawdataTE = getrawdata(wd, year, data_name='Quebec_Exports')
            rawdataTE_list.append(rawdataTE)
# add empty dataframes (na-filled) for years 2004-2008 of II (these will be estimated using the proportions of II in IC from subsequent years)
rawdataII = pd.DataFrame(data=None, index=range(41), columns=range(52))
for year in range(2004, 2008+1):
    yearindex = year - 2004
    rawdataII_list[yearindex] = rawdataII

# array for all sectors intermediate consumption
years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])
yearsColumns = np.arange(1995,2018+1)
# array for PBM array (2004-2008) before they are made into interieur arrays
# (imports subtracted with proportions from mid and new years)
asIC_PBM_Array = np.full(shape=(oldPBMNoSectors, oldPBMNoSectors, 5), fill_value=-1, dtype=float)
# array for already intérieur data (2009-2018)
asICArray = np.full(shape=(maxNoSectors, maxNoSectors, len(years)), fill_value=-1, dtype=float)
# array for already intérieur data for World (2004-2014)
#asICworldArray = np.full(shape=(56, 56, 11), fill_value=np.nan, dtype=float)
asICworldArray = asICworldArrayP
# array for already intérieur data for Rest of World (World minus Quebec) (2004-2014) (56 sectors)
# asICRoWArray = np.full(shape=(56, 56, 11), fill_value=np.nan, dtype=float)
# array for all sectors intermediate imports
asIIArray = np.full(shape=(maxNoSectors, maxNoSectors, len(years)), fill_value=-1, dtype=float)
# array for all sectors total production
asTPArray = np.full(shape=(maxNoSectors, len(years)), fill_value=-1, dtype=float)
# array for all sectors total production WORLD (2004-2014)
#asTPworldArray = np.full(shape=(56, 1, 11), fill_value=-1, dtype=float)
asTPworldArray = asTPworldArrayP
# array for all world sectors intermediate consumption
awsICwArray = np.full(shape=(35, 35, 2009 - 1995 + 1), fill_value=-1, dtype=float)
# array for all RoW (excluding quebec) sectors intermediate consumption
awsICRoWArray = np.full(shape=(35, 35, 2009 - 1995 + 1), fill_value=-1, dtype=float)
# array for all canada sectors intermediate exports and final demand
# (CAN has 56 sectors in IE, 5 categories in final demand)
#asIEFEcanArray = np.full(shape=(56, 2685 - 56 - 5, len(years)), fill_value=0, dtype=float)
asIEFEcanArray = asIEFEcanArrayP
# array for all canada sectors intermediate exports
asIEcanArray = np.full(shape=(56, 56, len(years)), fill_value=0, dtype=float)
# array for Canada final exports (added to a single column)
asFEcanArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors Québec total exports (final + intermediate)
asTEArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for final demand (added to a single column)
asFDArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors total production world
asTPwArray = np.full(shape=(56, len(years)), fill_value=-1, dtype=float)
# array for all sectors capital compensation (added to single column)
asCCArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors capital compensation (CAP) WORLD (added to single column)
asCCworldArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors labour compensation (LAB) (added to single column)
asLCArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors labour compensation WORLD (added to single column)
asLCworldArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors value added
asVAArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors world value added
#asVAworldArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
asVAworldArray = asVAworldArrayP
# array for all sectors Gross Fixed Capital Formation (added to single column)
asGFCFArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors Household demand (added to single column)
asHHArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors Government Expenditure (added to single column)
asGEArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors Inventories (added to single column)
asINVArray = np.full(shape=(maxNoSectors, len(years)), fill_value=0, dtype=float)
# array for all sectors World minus QC (initially this is World and then QC is subtracted) Final Demand
#asFDrowArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
asFDrowArray = asFDrowArrayP
# array for all sectors Gross Fixed Capital Formation (added to single column) WORLD
#asGFCFwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
asGFCFwArray = asGFCFwArrayP
# array for all sectors Household demand (added to single column) WORLD
#asHHwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
asHHwArray = asHHwArrayP
# array for all sectors Government Expenditure (added to single column) WORLD
#asGEwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
asGEwArray = asGEwArrayP
# array for all sectors Inventories (added to single column) WORLD
#asINVwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
asINVwArray = asINVwArrayP


# array for common sectors intermediate consumption (16 by 16 by years)
csICArray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), len(years)), fill_value=-1, dtype=float)
# array for common sectors intermediate imports (16 by 16 by years)
csIIArray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), len(years)), fill_value=-1, dtype=float)
# array for common sectors total production (16 by years)
csTPArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=-1, dtype=float)
# array for all world sectors intermediate consumption (WIOD2013 release, 1995-2009) (NOT USING THIS BECAUSE 1995-2003 data NOT AVAILABLE FOR QC)
#csICwArray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), 2009 - 1995 + 1), fill_value=-1,
#                    dtype=float)
# array for all world (WIOD2016 release, 2004-2014) sectors (56) intermediate consumption
csICworldArray = np.full(shape=(16, 16, 2014-2004+1), fill_value=-1,
                     dtype=float)
# array for already intérieur data for Rest of World (World minus Quebec) (2004-2014) (16 sectors)
csICRoWArray = np.full(shape=(16, 16, 2014-2004+1), fill_value=np.nan, dtype=float)
# array for RoW sectors intermediate consumption
#csICRoWArray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), 2009 - 1995 + 1), fill_value=-1,
#                     dtype=float)
# array for canada common sectors intermediate exports
csIEcanArray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), len(years)), fill_value=0,
                       dtype=float)
# array for canada common sectors intermediary exports proportions of total exports
csIEpcanArray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), len(years)), fill_value=0,
                        dtype=float)
# array for canada common sectors final exports
csFEcanArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for canada common sectors final exports proportions of total exports
csFEpcanArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for canada common sectors total exports
csTEcanArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for Québec common sectors total exports (final + intermediate)
csTEArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for Québec common sectors final exports
csFEArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for Québec common sectors intermediate exports
csIEArray = np.full(shape=(len(commonSectorsString), len(commonSectorsString), len(years)), fill_value=0,
                    dtype=float)
# array for common sectors total production World (WIOD2013, 1995-2009) (NOT USING THIS BECAUSE 1995-2003 DATA NOT AVAILABLE FOR QC)
#+csTPwArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# GOOD array for common sectors total production World (2004-2014; WIOD2016 release)
csTPworldArray = np.full(shape=(16, 11), fill_value=0, dtype=float)
# GOOD array for common sectors total production Rest of World (2004-2014; world-Quebec)
csTPRoWArray = np.full(shape=(16, 11), fill_value=0, dtype=float)
# array for common sectors capital compensation (added to single column)
# csCCArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors labour compensation (LAB) (added to single column)
#csLCArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors capital compensation (CAP) WORLD (added to single column)
csCCworldArrayTemp = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors labour compensation WORLD (added to single column)
csLCworldArrayTemp = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors value added (added to single column)
csVAArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for all common sectors world value added (added to single column)
csVAworldArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors Gross Fixed Capital Formation (added to single column)
csGFCFArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors Household demand (added to single column)
csHHArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors Government Expenditure (added to single column)
csGEArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors Inventories (added to single column)
csINVArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for Québec GDP (csCCArray + csLCArray + net tax + international trade margin (sp65 in EURegionalIOtable_2010))
GDPArray = np.full(shape=(1, len(years)), fill_value=0, dtype=float)
# array for common sectors World minus QC (initially this is World and then QC is subtracted) Final Demand
csFDrowArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors World Gross Fixed Capital Formation (added to single column)
csGFCFwArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors World Household demand (added to single column)
csHHwArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors World Government Expenditure (added to single column)
csGEwArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for common sectors World Inventories (added to single column)
csINVwArray = np.full(shape=(len(commonSectorsString), len(years)), fill_value=0, dtype=float)
# array for World GDP (csHHwArray + csGFCFwArray + csGEwArray + csINVwArray)
GDPwArray = np.full(shape=(1, len(years)), fill_value=0, dtype=float)



years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])
# get the Intérieur/PBM/Base IC data from 2009-2011 & 2014-2018
for year in years:
    setyeartype(year)
    print('year:', year)
    # here for 2012-2013 (no data) the dataframes are filled with nans which are kept in the array
    df_to_array(df_list=rawdataIC_list, data_name='IC', yr=year)  # creates asIC_PBMarray (2004-2008) & asICArray (2009-2018) CHECK
    df_to_array(df_list=rawdataIC_list, data_name='CC', yr=year)  # creates asCCArray CHECK
    df_to_array(df_list=rawdataIC_list, data_name='LC', yr=year)  # creates asLCarray CHECK
    df_to_array(df_list=rawdataIC_list, data_name='VA', yr=year)  # creates asVAarray CHECK
    df_to_array(df_list=rawdataIC_list, data_name='TP', yr=year)  # creates asTPArray CHECK
    df_to_array(df_list=rawdataBase_list, data_name='INV', yr=year)  # creates asINVArray CHECK
    df_to_array(df_list=rawdataBase_list, data_name='GE', yr=year)  # creates asGEArray CHECK
    df_to_array(df_list=rawdataBase_list, data_name='HH', yr=year)  # creates asHHArray CHECK
    df_to_array(df_list=rawdataBase_list, data_name='GFCF', yr=year)  # creates asGFCFArray CHECK
    df_to_array(df_list=rawdataII_list, data_name='II', yr=year)  # creates asIIArray CHECK
    df_to_array(df_list=rawdataTE_list, data_name='TE', yr=year)  # creates asTEArray CHECK
    # here for 2009 (no data) the dataframes are filled with nans which are kept in the array
    # df_to_array(df_list=rawdataTP_World_list, data_name='TP_World', yr=year)  # creates asTPworldArray CHECK
    # df_to_array(df_list=rawdataE_Canada_list, data_name='IE_CAN', yr=year)  # creates asIEFEcanArray CHECK
    # df_to_array(df_list=rawdataFD_World_list, data_name='FDw', yr=year)  # creates asFDrowArray CHECK


# FILL GAPS (2009 and 2012-2013) - time linearity assumption

# 2009
# asIEFEcanArray[:,:,5] = (asIEFEcanArray[:,:,4]+asIEFEcanArray[:,:,6])/2
# asFDrowArray[:,5] = (asFDrowArray[:,4]+asFDrowArray[:,6])/2



#2012-2013

d_temp = (asTPArray[:,10] - asTPArray[:,7])/3
asTPArray[:,8] = asTPArray[:,7] + d_temp
asTPArray[:,9] = asTPArray[:,7] + d_temp*2

d_temp = (asIIArray[:,:,10] - asIIArray[:,:,7])/3
asIIArray[:,:,8] = asIIArray[:,:,7] + d_temp
asIIArray[:,:,9] = asIIArray[:,:,7] + d_temp*2

d_temp = (asTEArray[:,10] - asTEArray[:,7])/3
asTEArray[:,8] = asTEArray[:,7] + d_temp
asTEArray[:,9] = asTEArray[:,7] + d_temp*2

d_temp = (asCCArray[:,10] - asCCArray[:,7])/3
asCCArray[:,8] = asCCArray[:,7] + d_temp
asCCArray[:,9] = asCCArray[:,7] + d_temp*2

d_temp = (asLCArray[:,10] - asLCArray[:,7])/3
asLCArray[:,8] = asLCArray[:,7] + d_temp
asLCArray[:,9] = asLCArray[:,7] + d_temp*2

d_temp = (asGFCFArray[:,10] - asGFCFArray[:,7])/3
asGFCFArray[:,8] = asGFCFArray[:,7] + d_temp
asGFCFArray[:,9] = asGFCFArray[:,7] + d_temp*2

d_temp = (asHHArray[:,10] - asHHArray[:,7])/3
asHHArray[:,8] = asHHArray[:,7] + d_temp
asHHArray[:,9] = asHHArray[:,7] + d_temp*2

d_temp = (asGEArray[:,10] - asGEArray[:,7])/3
asGEArray[:,8] = asGEArray[:,7] + d_temp
asGEArray[:,9] = asGEArray[:,7] + d_temp*2

d_temp = (asINVArray[:,10] - asINVArray[:,7])/3
asINVArray[:,8] = asINVArray[:,7] + d_temp
asINVArray[:,9] = asINVArray[:,7] + d_temp*2

d_temp = (asVAArray[:,10] - asVAArray[:,7])/3
asVAArray[:,8] = asVAArray[:,7] + d_temp
asVAArray[:,9] = asVAArray[:,7] + d_temp*2

# reaggregate from 56 --> 16 sectors (IC and TP)
world56_to_common = {1: [1, 2, 3], 2: [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23], 3: [4],
                     4: [24, 25], 5: [27], 6: [28], 7: [29], 8: [30], 9: [36], 10: [31, 32, 33, 34], 11: [35, 39],
                     12: [26, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 56], 13: [52], 14: [53], 15: [37, 38, 54],
                     16: [55]}

for year in range(2004, 2015):
    print('Reaggregating ICworld and TPworld from 56 to 16 sectors for year', year)
    yearIndex = year - 2004
    for csR in range(0, 16):
        row = world56_to_common[csR + 1]
        csTPworldArray[csR, yearIndex] = \
            np.sum(asTPworldArray[row - np.ones(len(row), dtype=int), :, yearIndex])
        for csC in range(0, 16):
            col = world56_to_common[csC + 1]
            csICworldArray[csR, csC, yearIndex] = getsums(asICworldArray, row, col, yearIndex)

#====================================================================================================



# aggregate asIEFE_Canada
# add up final exports columns
asFEcanArray = np.sum(asIEFEcanArray[:, 2408:2622, :], axis=1)  # needs 2009, 2015+

# aggregate IE array across countries
for i in range(0, 43):
    asIEcanArray = asIEcanArray + asIEFEcanArray[:, (i * 56):(55+1 + (56 * i)), :]

# ***** have to do these separately *****
years = np.array([2004, 2005, 2006, 2007, 2008])
# i.e. array by array (get_pre, then reagg, then freagg_check)
# because I was too lazy to create new variables in the functions
get_pre_rc_tots(years, asIC_PBM_Array)  # get col and row totals pre freagg (2004-2008)
asIC_PBM_Array = f_reagg(asIC_PBM_Array)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asIC_PBM_Array)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asTPArray)  # get col and row totals pre freagg (2004-2008)
asTPArray = f_reagg(asTPArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asTPArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asTEArray)  # get col and row totals pre freagg (2004-2008)
asTEArray = f_reagg(asTEArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asTEArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asCCArray)  # get col and row totals pre freagg (2004-2008)
asCCArray = f_reagg(asCCArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asCCArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asLCArray)  # get col and row totals pre freagg (2004-2008)
asLCArray = f_reagg(asLCArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asLCArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asVAArray)  # get col and row totals pre freagg (2004-2008)
asVAArray = f_reagg(asVAArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asVAArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asGEArray)  # get col and row totals pre freagg (2004-2008)
asGEArray = f_reagg(asGEArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asGEArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asHHArray)  # get col and row totals pre freagg (2004-2008)
asHHArray = f_reagg(asHHArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asHHArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asGFCFArray)  # get col and row totals pre freagg (2004-2008)
asGFCFArray = f_reagg(asGFCFArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asGFCFArray)  # get col and row totals post freagg and check if they match pre (2004-2008)

get_pre_rc_tots(years, asINVArray)  # get col and row totals pre freagg (2004-2008)
asINVArray = f_reagg(asINVArray)  # RE-AGG F1-F2-F3 (2004-2008)
freagg_check(years, asINVArray)  # get col and row totals post freagg and check if they match pre (2004-2008)


# put asIC_PBM_Array (2004-2008) into asICArray
asICArray[0:22, 0:22, 0:5] = asIC_PBM_Array

# fill 2012-2013 GAP (time linearity assumption)
d_temp = (asICArray[:,:,10] - asICArray[:,:,7])/3
asICArray[:,:,8] = asICArray[:,:,7] + d_temp
asICArray[:,:,9] = asICArray[:,:,7] + d_temp*2


# RE-AGG as -> cs
years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])

csICArray = as_to_cs(asICArray)
csIIArray = as_to_cs(asIIArray)
csTPArray = as_to_cs(asTPArray)
csTEArray = as_to_cs(asTEArray)
csCCArray = as_to_cs(asCCArray)
csLCArray = as_to_cs(asLCArray)
csGFCFArray = as_to_cs(asGFCFArray)
csHHArray = as_to_cs(asHHArray)
csGEArray = as_to_cs(asGEArray)
csINVArray = as_to_cs(asINVArray)
csVAArray = as_to_cs(asVAArray)



####


# reagg asIEcanArray & asFEcanArray (56 sectors) to csIEcanArray & csFEcanArray (16 common sectors) (see workplan data)
# reaggregate to 16*16
# world (56) -> common sectors (16)
world56_to_common = {
    1: [1, 2, 3],
    2: [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
    3: [4],
    4: [24, 25],
    5: [27],
    6: [28],
    7: [29],
    8: [30],
    9: [36],
    10: [31, 32, 33, 34],
    11: [35, 39],
    12: [26, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 56],
    13: [52],
    14: [53],
    15: [37, 38, 54],
    16: [55],
}

for year in years:
    yearIndex = getyearindex(years, year)
    for csR in range(0, 16):
        row = world56_to_common[csR + 1]
        csFEcanArray[csR, yearIndex] = \
            np.sum(asFEcanArray[row - np.ones(len(row), dtype=int), yearIndex])
        #csTPwArray[csR, yearIndex] = \
        #    np.sum(asTPwArray[row - np.ones(len(row), dtype=int), yearIndex])
        csFDrowArray[csR, yearIndex] = \
            np.sum(asFDrowArray[row - np.ones(len(row), dtype=int), yearIndex])
        for csC in range(0, 16):
            col = world56_to_common[csC + 1]
            csIEcanArray[csR, csC, yearIndex] = getsums(asIEcanArray, row, col, yearIndex)



# EXTRA STUFF (read & reagg)
# ----- GET WORLD CAP (as/csCCwArray - Capital Compensation) & LAB (as/csLCwArray - Labour Compensation) from Socio_Economic_Accounts_PIMPED.xlsx -----
# READ
stringname = 'Socio_Economic_Accounts_PIMPED.xlsx'
sheetname = 'DataInUSD1995'
skiprows = 0
index_col = list([0, 1, 2, 3])
rawdataCAPLABw = pd.read_excel(stringname, sheet_name=sheetname, index_col=index_col, header=0, skiprows=skiprows, usecols=None, nrows=None)
# aggregate across countries (43 countries, 168 rows per country)
rawdata = rawdataCAPLABw
rawdatar = rawdata.groupby(level=[1, 3]).sum() #aggregate across countries, retain sectors and CAP/LAB/VA

temp = rawdatar.loc['CAP']
tempar = temp.to_numpy().reshape(len(temp.index), len(temp.columns))
asCCworldArray[:, 0:len(temp.columns)] = tempar[:, :]  #asCCworldArray is 2000-2018 (2015-2018 empty)

temp = rawdatar.loc['LAB']
tempar = temp.to_numpy().reshape(len(temp.index), len(temp.columns))
asLCworldArray[:, 0:len(temp.columns)] = tempar[:, :] #asLCworldArray is 2000-2018 (2015-2018 empty)
# aggregate across common sectors
# reaggregate from 56 --> 16 sectors (CAPLABVA)
world56_to_common = {
    1: [1, 2, 3],
    2: [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
    3: [4],
    4: [24, 25],
    5: [27],
    6: [28],
    7: [29],
    8: [30],
    9: [36],
    10: [31, 32, 33, 34],
    11: [35, 39],
    12: [26, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 56],
    13: [52],
    14: [53],
    15: [37, 38, 54],
    16: [55],
}

for year in range(2000, 2014+1):
    print('Reaggregating CAP/LAB world from 56 to 16 sectors for year', year)
    yearIndex = year - 2000
    for csR in range(0, 16):
        row = world56_to_common[csR + 1]
        csCCworldArrayTemp[csR, yearIndex] = \
            np.sum(asCCworldArray[row - np.ones(len(row), dtype=int), yearIndex])
        csLCworldArrayTemp[csR, yearIndex] = \
            np.sum(asLCworldArray[row - np.ones(len(row), dtype=int), yearIndex])
# get % of VA for LAB and for CAP (assume these countries represent the whole world)
csCCworldArrayTemp/(csCCworldArrayTemp+csLCworldArrayTemp+1e-100)

# # ------GET World VA data from 2004-2014 (2009 skipped)------ (as and csVAworldArray have a small discrepancy less than 1% with original data not sure why)
# year_one_w = 2004
# for year in range(2004, 2014+1):
#     print('Reading raw data (VA world) for year', year)
#     if year == 2009:
#         pass
#     else:
#         yearIndex = year - year_one_w
#         stringname = 'WIOT_ROW/WIOT'+str(year)+'_Nov16_ROW.xlsb'
#         rawdataVAworld = pd.read_excel(stringname, sheet_name=str(year), index_col=None, header=list([2, 3, 4, 5]))
#         rawdata = rawdataVAworld.iloc[2469, 4:2465]  # take the VA row and columns
#         rawdatar = rawdata.groupby(level=[0]).sum()  # aggregate across countries, retain sectors and CAP/LAB/VA
#         tempar = rawdatar.to_numpy().reshape(len(rawdatar.index))
#         asVAworldArray[:, yearIndex] = rawdatar
#
# # FILL GAP (2009) (time linearity assumption)
# asVAworldArray[:,5] = (asVAworldArray[:,4]+asVAworldArray[:,6])/2

# aggregate across common sectors
# reaggregate from 56 --> 16 sectors (CAPLABVA)
world56_to_common = {1: [1, 2, 3], 2: [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23], 3: [4],
                     4: [24, 25], 5: [27], 6: [28], 7: [29], 8: [30], 9: [36], 10: [31, 32, 33, 34], 11: [35, 39],
                     12: [26, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 56], 13: [52], 14: [53], 15: [37, 38, 54],
                     16: [55]}

for year in range(2004, 2014+1):
    print('Reaggregating VA world from 56 to 16 sectors for year', year)
    yearIndex = year - 2004
    for csR in range(0, 16):
        row = world56_to_common[csR + 1]
        csVAworldArray[csR, yearIndex] = \
            np.sum(asVAworldArray[row - np.ones(len(row), dtype=int), yearIndex])

# get % of VA for LAB and for CAP (assume these countries represent the whole world)
# SO HERE hypothesis is that Capital Compensation and Labour Compensation proportions of Value Added
# taken from the countries in Socioeconomic_accounts_PIMPED.xlsx (many high income countries) are representative of the whole world. Could also use proportions from economy.xlsx 'World' tab.
csCCworldArray = csCCworldArrayTemp/(csLCworldArrayTemp+csCCworldArrayTemp+1e-100) * csVAworldArray
csLCworldArray = csLCworldArrayTemp/(csLCworldArrayTemp+csCCworldArrayTemp+1e-100) * csVAworldArray

# reaggregate from 56 --> 16 sectors (GFCF, GE, INV, HH)
world56_to_common = {1: [1, 2, 3], 2: [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23], 3: [4],
                     4: [24, 25], 5: [27], 6: [28], 7: [29], 8: [30], 9: [36], 10: [31, 32, 33, 34], 11: [35, 39],
                     12: [26, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 56], 13: [52], 14: [53], 15: [37, 38, 54],
                     16: [55]}

for year in range(2004, 2014+1):
    print('Reaggregating FDworld categories (GFCF, GE,INV, HH) from 56 to 16 sectors for year', year)
    yearIndex = year - 2004
    for csR in range(0, 16):
        row = world56_to_common[csR + 1]
        csGEwArray[csR, yearIndex] = \
            np.sum(asGEwArray[row - np.ones(len(row), dtype=int), yearIndex])
        csHHwArray[csR, yearIndex] = \
            np.sum(asHHwArray[row - np.ones(len(row), dtype=int), yearIndex])
        csGFCFwArray[csR, yearIndex] = \
            np.sum(asGFCFwArray[row - np.ones(len(row), dtype=int), yearIndex])
        csINVwArray[csR, yearIndex] = \
            np.sum(asINVwArray[row - np.ones(len(row), dtype=int), yearIndex])


# Get GDP
for year in range(2004, 2014+1):
    year_i = year - 2004
    # WORLD GDP ARRAY from (csGFCFwArray + csGEwArray + csINVwArray + csHHwArray)
    GDPwArray[:, year_i] = np.sum(csGFCFwArray[:,year_i], axis=0) + np.sum(csGEwArray[:,year_i], axis=0) +  np.sum(csINVwArray[:,year_i], axis=0) + np.sum(csHHwArray[:,year_i], axis=0)
    # QC GDP ARRAY
    GDPArray[:, year_i] = np.sum(csVAArray[:, year_i], axis=0)



#============================================================================================
# RoW = World - QC
csTPRoWArray[:, :] = csTPworldArray[:, :] - csTPArray[:, 0:11]  # 0:11 is 2004-2014 (csTPArray goes from 2004:2018, while csTPworldArray goes from 2004:2014)
csICRoWArray[:, :, :] = csICworldArray[:, :, :] - csICArray[:, :, 0:11]  # idem
#============================================================================================

# csTPwArray[:, :] = csTPwArray[:, :] - csTPArray[:, :]
# FDw - FDquebec = FDrow; FDquebec is HH+GFCF+GE+INV
csFDrowArray[:, :] = csFDrowArray[:, :] - csHHArray[:, :] - csGFCFArray[:, :] - csGEArray[:, :] - csINVArray[:, :]

# remove negative values (minus ones)
asICArray[asICArray[:, :, :] < 0] = 0
csICArray[csICArray[:, :, :] < 0] = 0
asIIArray[asIIArray[:, :, :] < 0] = 0
csIIArray[csIIArray[:, :, :] < 0] = 0
asTPArray[asTPArray[:, :] < 0] = 0
asTPwArray[asTPwArray[:, :] < 0] = 0
csTPArray[csTPArray[:, :] < 0] = 0
asIEcanArray[asIEcanArray[:, :, :] < 0] = 0
csIEcanArray[csIEcanArray[:, :, :] < 0] = 0
asFEcanArray[asFEcanArray[:, :] < 0] = 0
csFEcanArray[csFEcanArray[:, :] < 0] = 0
#csTPwArray[csTPwArray[:, :] < 0] = 0

# check that as -> cs aggregation is correct
checksums_as_cs(asICArray, csICArray)
checksums_as_cs(asIIArray, csIIArray)
checksums_as_cs(asTPArray, csTPArray)

checksums_as_cs(asIEcanArray, csIEcanArray)
checksums_as_cs(asFEcanArray, csFEcanArray)
checksums_as_cs(asTEArray, csTEArray)

# calculate the average and sd of the props of II and IC, to see if average can be used for 2004-2008
II_IC_prop = (csIIArray / (csICArray + csIIArray+1e-100))
# special case for the II_IC_prop[2,0,7] = infinity, let it be equal to the average of subsequent props (as for 2004-2008 years)
II_IC_prop[2, 0, 7] = sum(II_IC_prop[2, 0, [5, 6, 10, 11, 12, 13, 14]]) / 7
II_IC_prop_avg = np.full(shape=(II_IC_prop.shape[0], II_IC_prop.shape[1]), fill_value=-1, dtype=float)
# II_IC_prop_sd = np.full(shape=(II_IC_prop.shape[0],II_IC_prop.shape[1]), fill_value=-1, dtype=float)
proplen = 2018 - 2009 + 1 #THIS IS THE CORRECT LINE OF CODE BUT DOES NOT WORK BECAUSE OF 2012 data not reading error
#proplen = 8 # THIS IS NOT THE CORRECT LINE OF CODE
for row in range(0, II_IC_prop.shape[0]):
    for col in range(0, II_IC_prop.shape[1]):
        II_IC_prop_avg[row, col] = np.sum(II_IC_prop[row, col, [5, 6, 10, 11, 12, 13, 14]]) / 7 #THIS IS THE CORRECT LINE OF CODE BUT DOES NOT WORK BECAUSE OF 2012 data not reading error
        #II_IC_prop_avg[row, col] = np.sum(II_IC_prop) / proplen # THIS IS NOT THE CORRECT CODE
#       II_IC_prop_sd[row,col] = get_sd(II_IC_prop[row,col,:])

# II_IC_prop_cv = np.full(shape=(II_IC_prop.shape[0],II_IC_prop.shape[1]), fill_value=-1, dtype=float)
# II_IC_prop_cv = II_IC_prop_sd/II_IC_prop_avg*100

# csIC & csII arrays for 2004-2008
#   II+IC (2009-2018 csIC/IIArray) = Base (2004-2008 csICArray)
#   II/IC = prop ; IC = II/prop
#   Base - IC = II ; Base - II/prop = II --> II = II = Base / (1/prop + 1)

# totalscsICArray as a temporary duplicate of csICArray to be used for the
totalcsICArray = np.full(shape=(csICArray.shape[0], csICArray.shape[1], csICArray.shape[2]), fill_value=-1, dtype=float)
totalcsICArray[:, :, :] = csICArray[:, :, :]

for year in range(0, 5):
    csIIArray[:, :, year] = totalcsICArray[:, :, year] * II_IC_prop_avg[:, :]
    csICArray[:, :, year] = totalcsICArray[:, :, year] * (np.ones((16, 16)) - II_IC_prop_avg[:, :])

# check that it works (difference below should be 0 or very close)
checksums_ICtot_IC_II_reagg(totalcsICArray, csICArray, csIIArray)

# disaggregate csTEArray into csFEArray and csIEArray, using csFEcanArray and csIEcanArray
csTEcanArray[:, :] = csFEcanArray[:, :] + np.sum(csIEcanArray[:, :, :], axis=1)

csFEpcanArray[:, :] = csFEcanArray[:, :] / (csTEcanArray[:, :] + 1e-100)
csFEArray[:, :] = csTEArray[:, :] * csFEpcanArray[:, :]  # this looks fine but double check it gives the right values

for sector in range(0, 16):
    csIEpcanArray[sector, :, :] = csIEcanArray[sector, :, :] / (csTEcanArray[sector, :] + 1e-100)
    csIEArray[sector, :, :] = csTEArray[sector, :] * csIEpcanArray[sector, :, :]

csIEArray[15, :, :] = csTEArray[15, :] / 17  # the can proportions for sector 16 are nan, so split equally the very small values
csFEArray[15, :] = csTEArray[15, :]/17

# A matrices for IC (top left), II (bottom left), ICRoW (bottom right) and IE (top right) (see big book p.54)
A_IC = get_A(csICArray, csTPArray)  # 2004-2011 & 2014-2018
A_II = get_A(csIIArray, csTPArray)  # 2004-2011 & 2014-2018

years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014])
A_ICRoW = get_A(csICRoWArray, csTPRoWArray)  # this is for 2004-2014

A_IE = get_A(csIEArray, csTPRoWArray)
# A matrix for world model
A_ICw = get_A(csICworldArray, csTPworldArray)  #2004-2014, 2009 missing

# Combine the 4 A matrices into one to fit the format in economy.xlsx
temptop = np.concatenate((A_IC[:, :, 0:11], A_IE), axis=1)
tempbot = np.concatenate((A_II[:, :, 0:11], A_ICRoW), axis=1)
A4 = np.concatenate((temptop, tempbot), axis=0)


#=======================================================================================================================
# GETTING WORLD ENERGY INTENSITIES 16 sectors

# EI = TFE / TP
# Energy intensity (sector, source) = Total final energy (energy unit) (sector, source) / Total Production (TP) (monetary unit) (sector))
# IC matrix (sector by sector) = A matrix (sector by sector) * TP vector (sector)

#-1 Import:
#   WORLD TOTAL FINAL ENERGY (TFE) BY SECTOR (56) and SOURCE ('Emission relevant energy use' from Environmental Accounts of WIOD 2016 release)
#   Read 'per-country' files
#   Read <Emission relevant energy use> files for each country

# get list of all files in directory
os.chdir('/Users/Etguer/pymedeasQCdata/QC_IOT_data/QC_IOTpy_input')  # need to use '/' or '//' instead of '\'
directory = 'EmRel10/'
allfiles = [f for f in listdir(directory) if isfile(join(directory, f))]
years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014])
sumdfsList = list()
dfsourcedList = list()
for year in years:
    yearIndex = year-2004
    sumdfs = pd.read_excel(directory+allfiles[0], sheet_name=str(year), header = 0, index_col = 0) # create initial dataframe with the desired structure
    for file in (allfiles):
        if file == allfiles[0]:
            pass
        else:
            data = pd.read_excel(directory+file, sheet_name=str(year), header = 0, index_col = 0)
            sumdfs = sumdfs + data #aggregate into a single WORLD region
    sumdfsList.append(sumdfs)
    #   aggregate sources into 5 sources
    dfsourced = pd.DataFrame(data=None, index = sumdfs.index, columns = ['ELEC','HEAT','LIQUIDS','GAS','SOLIDS']) # create dataframe with MEDEAS-desired 5 energy sources
    dfsourced['HEAT'] = sumdfs[['ELECTR_HEATPROD', 'OTHSOURC']].sum(axis=1)
    dfsourced['ELEC'] = sumdfs['RENEWABLES_NUCLEAR']
    dfsourced['SOLIDS'] = sumdfs['COAL_COKE_CRUDE'] + sumdfs['WASTE'] / 2
    dfsourced['LIQUIDS'] = sumdfs[['JETFUEL', 'DIESEL', 'GASOLINE', 'FUEL_OIL', 'OTHPETRO', 'LIQUID_GASEOUS_BIOFUELS']].sum(axis=1) + sumdfs['WASTE'] / 2
    dfsourced['GAS'] = sumdfs[['NATGAS', 'OTHGAS']].sum(axis=1)
    dfsourcedList.append(dfsourced)
    #test sums for errors
    sumdfs.drop('TOTAL',axis=1).sum().sum() - dfsourced.sum().sum()

# array for all sectors world emission relevant energy use (final energy use, for energy intensities)
asFEUworldArray = np.full(fill_value = None, shape=(57,5,11))
# array for Households sector world emission relevant energy use (FEU HH, for energy intensities)
FEU_HHworldArray = np.full(fill_value = None, shape=(1,5,11))
for year in range(2004,2015):
    yearIndex = year-2004
# convert dataframe to array and remove last row which are not relevant (Total energy across all sectors).
    asFEUworldArray[:,:,yearIndex] = np.delete(dfsourcedList[yearIndex].to_numpy(),obj=[-2],axis=0)
# keep last row (Household FEU)
    #FEU_HHworldArray[:,:,yearIndex] = dfsourcedList[yearIndex].to_numpy()[-1,:]
# convert TJ (original data) to EJ (for MEDEAS input)
asFEUworldArray = asFEUworldArray * 10e-6
#FEU_HHworldArray = FEU_HHworldArray * 10e-6
#aggregate rows (56 sectors to 16 common sectors)
world56_to_common = {1: [1, 2, 3], 2: [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23], 3: [4],
                     4: [24, 25], 5: [27], 6: [28], 7: [29], 8: [30], 9: [36], 10: [31, 32, 33, 34], 11: [35, 39],
                     12: [26, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 56], 13: [52], 14: [53], 15: [37, 38, 54],
                     16: [55]}

csFEUworldArray = np.full(fill_value=None,shape=(17,5,11))
for year in range(2004, 2015):
    print('Reaggregating asFEUworldArray from 56 to 16 sectors for year', year)
    yearIndex = year - 2004
    for csR in range(0, 16):
        row = world56_to_common[csR + 1]
        csFEUworldArray[csR, :, yearIndex] = np.sum(asFEUworldArray[row - np.ones(len(row), dtype=int), :, yearIndex], axis=0) #add 56sectors to make 16sectors according to dictionary

# Add HH sector (same as for 'as')
csFEUworldArray[-1,:,:] = asFEUworldArray[-1,:,:]

#test that sum is same from 'as' to 'cs'
np.sum(asFEUworldArray)- np.sum(csFEUworldArray) # it works they're the same


#-3 Calculate World energy intensities (16 sectors (+ Households HH), 5 sources) as TFE / TP (TFE[HH] / sum acrpss sectors of HH for households)
# each sector (16) by source (5) : EI (16,5,11) = TFE(16,5,11)/TP(16,11)
EIworldArray = np.full(fill_value=None,shape=(17,11,5)) #(sectors,years,sources)
#prepare arrays by giving same shape as receiving array
csFEUworldArray = np.swapaxes(csFEUworldArray, 1, 2)

for source in range(0,5):
    EIworldArray[1:17,:,source] = csFEUworldArray[0:16,:,source]/csTPworldArray
    EIworldArray[0,:,source] = csFEUworldArray[16,:,source]/np.sum(csHHwArray[:,0:11],axis=0) #add HH sector

#Make array 2D with ELEC, HEAT, LIQUIDS, GAS, SOLIDS as consecutive rows (sectors, years)
EIworldMedeasFormat = np.full(fill_value=None, shape=(17*5,11))
for source in range(0,5):
    rowstart = source*17
    rowend = (source+1)*17
    EIworldMedeasFormat[rowstart:rowend, :] = EIworldArray[:,:,source]

sectors = np.tile(np.concatenate((['Households'], commonSectorsString)),5)
sources = np.repeat(['ELEC','HEAT','LIQUIDS','GAS','SOLIDS'],17)

EIindex = pd.MultiIndex.from_arrays(arrays = [sectors, sources], names = ['Energy Intensities', '(EJ/M$)'])

# At a quick glance it seems correct, values are within same oder of magnitude as the MEDEAS world energy intensities for 14 sectors

#=======================================================================================================================
#GETTING QC ENERGY INTENSITIES
#read energy use QC for energy intensities
path = 'energyusepersectorQCcompilation2024.xlsx'
tempdf = pd.read_excel(path, sheet_name='energy_use_sector_year_PJ', index_col=[0,1],
                       header=0)
tempar = tempdf.to_numpy().reshape(len(tempdf.index), len(tempdf.columns))
FEU_QCArray = tempar/1000 #convert read data (PJ) to EJ
#calculate energy intensititties
#all sectors except HH: EI (sector, source) = FEU (sector, source) / TP (sector)
#HH sector:             EI (source)         = FEU[HH sector] (source) / HH (sector)
EI_QCArray = np.full(fill_value=np.nan, shape=(FEU_QCArray.shape))

for row in range(0,5):
    startrowTP=row+1+16*row
    endrowTP=startrowTP+16
    rowHH=startrowTP-1
    EI_QCArray[startrowTP:endrowTP,9:] = FEU_QCArray[startrowTP:endrowTP,9:]/csTPArray
    EI_QCArray[rowHH,9:] = FEU_QCArray[rowHH,9:]/np.sum(csHHArray,axis=0)

#=======================================================================================================================
#REGRESSIONS FOR ECONOMIC COMPUTATIONS
#mean intensity rates





# Read all Mean Intensity Rates and Regressions coefficients and write them to economydata.xlsx
# Read them
os.chdir('/Users/Etguer/pymedeasQCdata/QC_IOT_data/')  # need to use '/' or '//' instead of '\'
directoryCoeffs = 'jupyter_output/'
writepath = os.getcwd() + '/' + 'economydata.xlsx'
#list of names
coeffFileNames = ['exp_qc_gdp.csv', 'gfcf_qc_cap.csv','hd_qc_lab.csv','meir_qc.csv',
                  'gfcf_w_cap.csv','hd_w_lab.csv','meir_w.csv']
#list of sheet name for each
coeffSheetNames = ['Quebec', 'Quebec', 'Quebec', 'Quebec',
                   'World', 'World', 'World']
#list of column names for each
coeffColumnNames = [['beta 0 (indep)','beta 1 (GDP W)'], ['beta 0 (indep)', 'beta 1 (cap C)'], ['beta 0 (indep)', 'beta 1 (Lab C)'],
                    ['Mean rate'], ['beta 0 (indep)', 'beta 1 (cap C)'], ['beta 0 (indep)', 'beta 1 (Lab C)'], ['Mean rate']]
#list of [startrow,startcol] for each
coeffExcelLocation = [[128,28],[38,28],[56,28],[146,28],  # locations of first datapoint in excel
                      [38,28],[56,28],[111,28]]




# Write and include headers along
#


#TO DO: clean up code to make it elegant and organized

#=======================================================================================================================
#=======================================================================================================================

# 1- add index and headers to all dataframe
#empty (M$) column
emptyMindex = np.repeat([''], 16)
#sectors names in 1st column as index
sectorsIndex = commonSectorsString
# column names in 1st row as header
yearsColumns = np.arange(1995, 2019)

#Quebec stuff for writing in excel spreadsheet in the MEDEAS input format

#econometric variables
OutputQCIndexListNames = [
    'Capital Compensation',
    'Labour Compensation',
    'Gross Fixed Capital Formation',
    'Households demand',
    'Government expenditures',
    'Change in inventories',
    'Final demand RoW',
    'Exports of FINALS GOODS to RoW',
]

OutputQCIndexListNamesShort = ['CC', 'LC', 'GFCF', 'HH', 'GE', 'INV', 'FDrow', 'FErow']

#multiindex because the first two columns are indices
OutputQCIndexList = list()
for name in OutputQCIndexListNames:
    temp = pd.MultiIndex.from_arrays(arrays = ([sectorsIndex, emptyMindex]), names = [name, '(M$)'])
    OutputQCIndexList.append(temp)
#list all the arrays to for loop it swiftly
OutputQCValuesList = [csCCArray, csLCArray, csGFCFArray, csHHArray,csGEArray, csINVArray, csFDrowArray, csFEArray]
#Excel locations QC
OutputQuebecExcelLocationList = [ #[row, col] for each econometric variable
    [1,1],[19,1],[37,1],[55,1],[73,1],[91,1],[109,1],[127, 1], # CC, LC, ..., FErow
    [146,1], #EI
    [233,1],[235,1], #GDP
    [239,1] #A matrices
]

#World stuff for writing in excel spreadsheet in the MEDEAS input format
OutputWorldIndexListNames = OutputQCIndexListNames[0:-2] #remove last two items
OutputWorldIndexListNmesShort = OutputQCIndexListNamesShort[0:-2]
OutputWorldIndexList = OutputQCIndexList[0:-2]
OutputWorldValuesList =  [csCCworldArray, csLCworldArray, csGFCFwArray, csHHwArray, csGEwArray, csINVwArray]
OutputWorldExcelLocationList = [ #[row, col] for each econometric variable
    [1,1],[19,1],[37,1],[55,1],[73,1],[91,1],# CC, LC, ..., INV
    [110,1], #EI
    [197,1],[199,1], #GDP
    [203,1] #A matrices
]

# WRITE all data to an excel workbook (economydata.xlsx)
path = '/Users/Etguer/pymedeasQCdata/QC_IOT_data/medeasQC_input/economydata.xlsx'

# PREPARE arrays to be made into dataframes, create (multi)indices

# EXTEND arrays for missing years
#make extensions
#empty array for 1995-2003 / 2015-2018 years for econometric data and EI
csEmptyYearsArray = np.full(fill_value = np.nan, shape = (16, len(np.arange(1995, 2004))))
EIemptyYearsArray1 = np.full(fill_value = np.nan, shape = (EIworldMedeasFormat.shape[0], len(np.arange(1995,2004))))
EIemptyYearsArray2 = np.full(fill_value = np.nan, shape = (EIworldMedeasFormat.shape[0], len(np.arange(2014,2018))))
#empty array for 1995-2003 years for GDP data
GDPEmptyYearsArray = np.full(fill_value = np.nan, shape = (1, len(np.arange(1995, 2004))))
#concatenate into single array
GDPArrayMedeasFormat = np.concatenate((GDPEmptyYearsArray, GDPArray), axis = 1)
GDPwArrayMedeasFormat = np.concatenate((GDPEmptyYearsArray, GDPwArray), axis = 1)
#empty arrays for 1995-2003 / 2015-2018 years for A matrices
#QC
A4emptyYearsArray1 = np.full(fill_value = np.nan, shape = np.concatenate((A4[:,:,0].shape,[len(np.arange(1995,2004))])))
A4emptyYearsArray2 = np.full(fill_value = np.nan, shape = np.concatenate((A4[:,:,0].shape,[len(np.arange(2014,2018))])))
#World
A_ICwEmptyYearsArray1 = np.full(fill_value = np.nan, shape = np.concatenate((A_ICw[:,:,0].shape,[len(np.arange(1995,2004))])))
A_ICwEmptyYearsArray2 = np.full(fill_value = np.nan, shape = np.concatenate((A_ICw[:,:,0].shape,[len(np.arange(2014,2018))])))
#concatenate into single arrays
A4medeasFormat = np.concatenate((A4emptyYearsArray1,A4,A4emptyYearsArray2), axis=2)
A_ICwMedeasFormat = np.concatenate((A_ICwEmptyYearsArray1,A_ICw,A_ICwEmptyYearsArray2), axis=2)

#create index/columns for A matrices
#QC
a=['QC ']*len(csString)+['RoW ']*len(csString)
b=csString*2
A4QCcolind = [i + j for i, j in zip(a,b)]
#World
a=['World ']*len(csString)
b=csString
A4Wcolind = [i + j for i, j in zip(a,b)]

with pd.ExcelWriter(path, engine='openpyxl') as writer:
# WRITE QC econometric variables (CC, LC, GFCF, ..., FErow)
    for i in range(len(OutputQCIndexListNames)):
        data = pd.DataFrame(
            data=np.append(csEmptyYearsArray, OutputQCValuesList[i], axis=1),
            columns=yearsColumns,
            index=OutputQCIndexList[i]
        )
        startrow = OutputQuebecExcelLocationList[i][0] - 1
        startcol = OutputQuebecExcelLocationList[i][1] - 1
        sheetname='Quebec'
        data.to_excel(
            writer,
            sheet_name=sheetname,
            na_rep='na',
            header=True,
            index=True,
            startrow=startrow,
            startcol=startcol
        )
        add_defined_name_section(writer, sheetname, "TODO", data, startrow, startcol, header=True, index=True)
# WRITE World econometric variables (CC, LC, GFCF, ..., INV)
    for i in range(len(OutputWorldIndexListNames)):
        data = pd.DataFrame(
            data=np.append(csEmptyYearsArray, OutputWorldValuesList[i], axis=1),
            columns=yearsColumns,
            index=OutputWorldIndexList[i]
        )
        startrow=OutputWorldExcelLocationList[i][0]-1
        startcol=OutputWorldExcelLocationList[i][1]-1
        sheetname='World'
        data.to_excel(
            writer,
            sheet_name=sheetname,
            na_rep='na',
            header=True,
            index=True,
            startrow=startrow,
            startcol=startcol
        )
        add_defined_name_section(writer, sheetname, "TODO", data, startrow, startcol, header=True, index=True)
# WRITE EI data
    # QC
    data = pd.DataFrame(data=EI_QCArray, index=EIindex, columns=yearsColumns)
    data.to_excel(writer, sheet_name='Quebec', na_rep='na', header=True, index=True, startrow=145-1, startcol=1-1)
    add_defined_name_section(writer, 'Quebec', "TODO", data, 145-1, 1-1, header=True, index=True)

    #WORLD
    data = pd.DataFrame(
        data=np.concatenate((EIemptyYearsArray1, EIworldMedeasFormat, EIemptyYearsArray2,), axis=1),
        index=EIindex,
        columns=yearsColumns
    )
    data.to_excel(writer, sheet_name='World', na_rep='na', header=True, index=True, startrow=110-1, startcol=1-1)
    add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=True)

#GDP for both
    #QC
    data = pd.DataFrame(data=GDPArrayMedeasFormat, columns=yearsColumns)
    data.to_excel(writer, sheet_name='Quebec', na_rep='na', header=True, index=False, startrow=233-1, startcol=3-1)
    add_defined_name_section(writer, 'Quebec', "TODO", data, 145-1, 1-1, header=True, index=True)

    data = pd.DataFrame(data=None, columns=['historic GDP', '(M$)'])
    data.to_excel(writer, sheet_name='Quebec', index=False, header=True, startrow=233-1,startcol=1-1)
    add_defined_name_section(writer, 'Quebec', "TODO", data, 145-1, 1-1, header=True, index=True)
    #World
    data = pd.DataFrame(data=GDPwArrayMedeasFormat, columns=yearsColumns)
    data.to_excel(writer, sheet_name='World', na_rep='na', header=True, index=False, startrow=197-1, startcol=3-1)
    add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=True)

    data = pd.DataFrame(data=None, columns=['historic GDP', '(M$)'])
    data.to_excel(writer, sheet_name='World', header=True, index=False, startrow=197 - 1, startcol=1 - 1)
    add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=True)
#GDPpc projection growth for both
    #QC
    data = pd.DataFrame(data=np.full(fill_value=np.nan, shape=(1,24)), columns=yearsColumns)
    data.to_excel(writer, sheet_name='Quebec', na_rep='na', header=True, index=False, startrow=235 - 1, startcol=3 - 1)
    add_defined_name_section(writer, 'Quebec', "TODO", data, 145-1, 1-1, header=True, index=True)
    
    data = pd.DataFrame(data=None, columns=['GDPpc projection growth', '(Dmnl)'])
    data.to_excel(writer, sheet_name='Quebec', header=True, index=False, startrow=235 - 1, startcol=1 - 1)
    add_defined_name_section(writer, 'Quebec', "TODO", data, 145-1, 1-1, header=True, index=True)
    
    #World
    data = pd.DataFrame(data=np.full(fill_value=np.nan, shape=(1, 24)), columns=yearsColumns)
    data.to_excel(writer, sheet_name='World', na_rep='na', header=True, index=False, startrow=199 - 1, startcol=3 - 1)
    add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=True)


    data = pd.DataFrame(data=None, columns=['GDPpc projection growth', '(Dmnl)'])
    data.to_excel(writer, sheet_name='World', header=True, index=False, startrow=199 - 1, startcol=1 - 1)
    add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=True)

    # WRITE A matrices
# Write header of the A matrix section
    data = pd.DataFrame(data=None, columns=['A matrix', '(Dmnl)'])
    data.to_excel(writer, sheet_name='Quebec', index=None, startrow= 238-1, startcol=1-1)
    add_defined_name_section(writer, 'Quebec', "TODO", data, 145-1, 1-1, header=True, index=True)

    data = pd.DataFrame(data=None, columns=['A matrix', '(Dmnl)'])
    data.to_excel(writer, sheet_name='World', index=None, startrow= 202-1, startcol=1-1)
    add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=True)
# Write A matrices for QC and World
    for year in yearsColumns:
        yearIndex = year - 1995
# A4 matrix for QC
        startrow = 239-1+yearIndex*33  #239 is excel actual row number, 33 is gap between consecutive years
        data = pd.DataFrame(data=A4medeasFormat[:, :, yearIndex], columns=A4QCcolind, index=A4QCcolind)
        data.index.name = str(year)
        data.to_excel(writer, sheet_name='Quebec', na_rep='na', header=True, startrow=startrow, startcol=1-1)
        add_defined_name_section(writer, 'Quebec', "TODO", data, 145-1, 1-1, header=True, index=True)
# A matrix for World
        startrow = 203-1+yearIndex*17 #203 is excel actual row number, 17 is gap between consecutive years
        data = pd.DataFrame(data=A_ICwMedeasFormat[:, :, yearIndex], columns=A4Wcolind, index=A4Wcolind)
        data.index.name = str(year)
        data.to_excel(writer, sheet_name='World', na_rep='na', header=True, startrow=startrow, startcol=1-1)
        add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=True)


#Coefficients (econometric regressions and energy intensity rates from jupyter notebooks)
# Read
    for i in range(len(coeffFileNames)):
        data = pd.read_csv(directoryCoeffs + coeffFileNames[i], header=None, index_col=None)
        #drop 3rd columns when applicable
        if data.shape[1] > 2:
            data.drop(labels=2, axis=1, inplace=True)
        data.columns = coeffColumnNames[i]
        #write in economydata.xlsx
        data.to_excel(writer, sheet_name=coeffSheetNames[i], na_rep='na', header=True,
                      index=False, startrow=coeffExcelLocation[i][0]-1-1 # -1  because there is a header, -1 because first row is 0
                      , startcol=coeffExcelLocation[i][1]-1) # -1 because first row is 0
        add_defined_name_section(writer, 'World', "TODO", data, 145-1, 1-1, header=True, index=False)

#TO DO: clean up code to make it elegant and organized

# might need to change engine to xlsxwriter (instead of openpyxl),
# see https://stackoverflow.com/questions/51531715/pandas-dataframe-to-excel-with-defined-name-range

#GET DEFINED NAMES FROM economy.lsx original file, output them in console and then copy paste them in this .py file
# os.chdir('/Users/Etguer/pymedeasQCdata/QC_IOT_data/medeasQC_input/Original input files')
# xlsxfile = 'economyOG.xlsx'
# wb = openpyxl.load_workbook(xlsxfile)
# ws = wb['World']
# [[value.name, value.attr_text] for value in ws.defined_names.values() ]

defnsWorld_List = [ #all defined names edited to my custom Quebec/World economy.xlsx data
 ['beta_0_GFCF', '$AB$38:$AB$53'], ['beta_0_HD', '$AB$56:$AB$71'],
 ['beta_1_GFCF', '$AC$38:$AC$38'], ['beta_1_HD', '$AC$56:$AC$56'],

 ['historic_A_Matrix_year1995', '$B$204:$Q$219'], ['historic_A_Matrix_year1996', '$B$221:$Q$236'],
 ['historic_A_Matrix_year1997', '$B$238:$Q$253'], ['historic_A_Matrix_year1998', '$B$255:$Q$270'],
 ['historic_A_Matrix_year1999', '$B$272:$Q$287'],
 ['historic_A_Matrix_year2000', '$B$289:$Q$304'], ['historic_A_Matrix_year2001', '$B$306:$Q$321'],
 ['historic_A_Matrix_year2002', '$B$323:$Q$338'], ['historic_A_Matrix_year2003', '$B$340:$Q$355'],
 ['historic_A_Matrix_year2004', '$B$357:$Q$372'], ['historic_A_Matrix_year2005', '$B$374:$Q$389'],
 ['historic_A_Matrix_year2006', '$B$391:$Q$406'], ['historic_A_Matrix_year2007', '$B$408:$Q$423'],
 ['historic_A_Matrix_year2008', '$B$425:$Q$440'], ['historic_A_Matrix_year2009', '$B$442:$Q$457'],
 ['historic_A_Matrix_year2010', '$B$459:$Q$474'], ['historic_A_Matrix_year2011', '$B$476:$Q$491'],
 ['historic_A_Matrix_year2012', '$B$493:$Q$508'], ['historic_A_Matrix_year2013', '$B$510:$Q$525'],
 ['historic_A_Matrix_year2014', '$B$527:$Q$542'],

 ['historic_capital_compensation', '$C$2:$V$17'], #setting it as in OG economy file, that is to say all sectors from 1995 to ***2014***
 ['historic_change_in_inventories', '$C$92:$Q$107'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_GFCF', '$C$38:$Q$53'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_goverment_expenditures', '$C$74:$Q$89'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_HD', '$C$56:$Q$71'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_labour_compensation', '$C$20:$V$35'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009

 ['historic_final_energy_intensity_electricity', '$C$111:$Q$127'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_final_energy_intensity_gases', '$C$162:$Q$178'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_final_energy_intensity_heat', '$C$128:$Q$144'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_final_energy_intensity_liquids', '$C$145:$Q$161'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009
 ['historic_final_energy_intensity_solids', '$C$179:$Q$195'], #setting it as in OG economy file, that is to say all sectors from 1995 to 2009

 ['historic_GDP', '$C$198:$V$198'],

 ['historic_mean_rate_energy_intensity_electricity', '$AB$111:$AB$127'],
 ['historic_mean_rate_energy_intensity_gases', '$AB$162:$AB$178'],
 ['historic_mean_rate_energy_intensity_heat', '$AB$128:$AB$144'],
 ['historic_mean_rate_energy_intensity_liquids', '$AB$145:$AB$161'],
 ['historic_mean_rate_energy_intensity_solids', '$AB$179:$AB$195'],

 ['input_GDPpc_annual_growth', '$C$200:$AL$200'], ['time_index_projection', '$C$199:$AL$199'], #this data be missin in my economy file
 ['time_index2009', '$C$1:$Q$1'], ['time_index2014', '$C$1:$V$1']
 ]

#open economydata.xlsx, add defined names to World sheet and save
os.chdir('/Users/Etguer/pymedeasQCdata/QC_IOT_data/medeasQC_input/QCinput2024')  # need to use '/' or '//' instead of '\'
xlsxfile = 'economydata.xlsx'
wb = openpyxl.load_workbook(xlsxfile)
ws = wb['World']
ws.title = 'World'
for element in defnsWorld_List:
    ref = f"{quote_sheetname(ws.title)}!{absolute_coordinate(element[1])}"
    defn = DefinedName(element[0], attr_text = ref)
    ws.defined_names.add(defn)

wb.save('economydatadefname.xlsx')

#TO DO:
# defined names for Quebec sheet
#GDP growth rates change header from 1995-... to from 2015 (?) - 2050 as in economyOG.xlsx
#when I run pymedeas2 there is a defined names bug - even with original catalonia model, so this might have to do with my installed version of openpyxl? or I dunno


# CHECKED asIC_PBM_Array into asICArray -> csICArray + checksums
# CHECKED asICArray -> csICArray (interieur) (IC = intermediate consumption) + checksums
# CHECKED asII arrays -> csII arrays (importations totales = importations internationales + importations interprovinciales) (II = intermediate imports) for 2009-2018
# CHECKED use proportion to define csIC & csII arrays for 2004-2008
# CHECKED asTP arrays -> csTP arrays for all years
# CHECKED RoW IC reaggregated to 16 common sectors
# CHECKED Intermediary exports and Final exports estimated from Canadian proportions and reaggregated to 16 common sectors
# CHECKED Get asTEArray (quebec total exports)
# CHECKED aggregate asTEArray to csTEArray
# CHECKED get Canada IE and FE proportions from csEEcanArray & csIEcanArray
# CHECKED apply proportions to csTEArray to obtain csIEArray & csFEArray
# CHECKED Exports of final goods to RoW   ***csFEArray***
# CHECKED get total output for RoW from WIOT data
# CHECKED get A matrix for IE (bottom right matrix) from csIEArray / csTPwArray
# CHECKED get Final Demand categories in asXXArray (HH demand, GFCF, GE, INV)
# CHECKED get Final Demand categories arrays into csXXArray (freagg(), then as_to_cs())
# CHECKED Final demand RoW : get from WIOT and reaggregate  ***csFDrowArray***
# CHECKED All four A matrices  (A4) written to excel, only works for 2004-2009 so far (see 2012-2013 reading data bug, and some stuff missing for 2014-2018 but I think it's a small bug somewhere)
# CHECKED Regression coefficients from Jupyter Notebook : get data into economy.xlsx, then run it with the jupyter script (need to tweak a few parameters such as years)
# CHECKED energy use per sector (data is in Etguer/Documents/QC_IOT_data/input/raw energy use per sector QC/, reagged using excel)
# **CHECKED** DONE WRONG historic GDP (I used historical GDP data but it should be done using A matrices, not sure how)
# CHECKED energy : intensities per sector and source (need total energy use per source per sector) (PRETTY SURE I HAVE THIS INFO IN OEE ENERGY USE PER SECTOR ON GDRIVE, DOING THE AGGREGATE BY HAND IN GOOGLE SHEETS WORKPLAN DATA)
# CHECKED Fix code to get all data for QC after 2008 (fill gaps, debug for 2012 if 2012 should be there, maybe not, review data) (file was corrupted or something, took back the original 2012 file)
# CHECKED WORLD economic data needs to be aggregated to 16 sectors, using most recent WIOD data would be better (and easier? or not). use code from above, only need one A matrix
# CHECKED (have to review for bugs) A matrix for world model: all countries IC including Québec (awsICwArray and csICwArray) currently I have it as taken from previous inputs_world.xlsx (good from 1995-2009),
# CHECKED(?)  otherwise for 2009+ would need to get the raw data, reagg and divide by TP_world as I did for QC
# CHECKED(?)  A matrix for RoW IC (bot-right) for Québec model: RoW IC (csICRoWArray) = World IC from A matrix above (csICwArray) - Québec IC (csICArray), so I need to get the world data anyways
# CHECKED(?)  for now I am using the same matrix csICwArray for both
# CHECKED (1) Get World IC raw data for 2004-2014 ; reaggregate to 16 sectors  (IN PROGRESS, see code above, variables need to be defined first)
# CHECKED (2) Make two arrays: one for World IC (csICworldArray), one for RoW IC (World minus Québec; csICRoWArray)
# CHECKED (3) Convert to A matrix by dividing IC/TP

# __________________________
# CHECK A matrix (A_ICw) (all IC, which includes exports and imports) -> from csICworldArray/csTPworldArray
#       historic GDP (World) from HH+GFCF+GE+INV
# __________________________
# FOR QUÉBEC and 16 sectors:
# CHECK csCCArray, csLCArray, csGFCFArray, csHHArray, csGEArray, csINVArray, csFDrowArray,
# CHECK csFEArray (APPARENTLY I HAVE THIS)
#       historic GDP (Québec) from CC+LC+nettaxes+international trade margin (for now using CC+LC+net taxes, can't figure if international trade margin is already included)
# CHECK A_IC
# CHECK A_II
# CHECK A_ICRoW
# CHECK A_IE
# __________________________


# =============================================================================
# TO DO

# 0- fix bugs throughout code
# CHECK (I limited the final outputs to 11 years when it created errors)  make all arrays go from year 2004-2014 (11 years), not 2004-2018 (15 years)

# CHECKED A_IC 2011 has no decimals
# ANSWER: raw data has no decimals

# 1- (CHECK ? -DOUBLE CHECK AFTER RERUN) FILL ALL GAPS WITH LINEAR ASSUMPTIONS FIRST (2009, 2012-2013)
#   BUT BEFORE CHECK THAT THE GAPS ARE LEGITIMATE (retrace them)

#   GAPS in Quebec A4 matrix (economyA4.xlsx) :
#   A_IE & A_ICw (2009); A_IC & A_II & A_IE (2012, 2013)

#   GAPS in Québec (economytest.xlsx):
#   CC (2012-2013); LC (2012-2013); GFCF (2012-2013); HH (2012-2013);
#   GE (2012-2013 & most sectors are suspicious 0s for all years); INV (2013-2013 & some sectors are suspicious 0s);
#   FDrow (2009,2012-2013); FE (2009, 2012, 2013 except Private_Households)

#   GAPS in WORLD (WIOT): 2009 (TP_World, FD_World, E_Canada)

#   estimate all missing years using linear model (2012-2013 IC QC and 2009 TE QC and I think the rest will follow)
#   except for 2015-2018 A_IE, this depends on IEcan which does not exist for 2015-2018,
#   so gotta check if exports is a constant proportion of total production or something)

# 2- SPLIT AGRICULTURE AND LOGGING
# Supporting activities half/half

# CHECK 3- FOR WORLD and 16 sectors:
#  CHECK   GET csCCworldArray, csLCworldArray from SocioEconomic Accounts.xlsx  (These are currently faked at 14 sectors divided in 16 (2 sectors are split in halves).
#  From Socio_Economic_Accounts_PIMPED.xlsx (the values are in 1995 national currency):
# -Get exchange rate for each currency in US1995$ and multiply (SocioEconomic_Accounts_PIMPED.xlsx)
# -Read in python
# -Aggregate across countries
# -Aggregate across common sectors
# -Get proportions of CAP & LAB vs VA for each sector
# -Assume these proportions apply to world? YEAH
# -Apply them to VA in the WIOT IO tables to get CAP and LAB time series for world
#  In socioecon. accounts VA = CAP + LAB (COMP is compensation for employees, LAB includes COMP + self-employed))

# CHECK 4- FOR WORLD also get (from WIOT_YEAR_Nov16_ROW.xlsx)
#        csGFCFwArray, csHHwArray, csGEwArray, csINVwArray (these are in my world data, rawdataFD_World_list, asFDrowArray) (These are also currently faked at 14 sectors divided in 16 (2 sectors are split in halves).
#        reagg them to 16 sectors (this is sum of all FD)

# CHECK 5- HISTORIC GPDs WORLD AND QUÉBEC (MONICA/FRANCISCA)
# CHECK Figure how to calculate them
#       WORLD: historic GDP from HH+GFCF+GE+INV (as per Monica's method in economy.xlsx[World])
#       QUÉBEC: historic GDP from CC+LC+nettaxes (+international trade margin - not used for now because I can't find it in my QC data) (defined below ctrl+_f) (as per MEDEAS D4.2 p.53, except for international trade margins)

# 6- BACKUPS/CHECKUPS
# write all my matrices outside, so I don't have to rerun this code forever
# write down detailed methodology (assumptions, data sources) for each dataset
#  -methods to cite for all the VA estimation stuff I did: Sources_SEA.pdf in QC_IOT_data (WIOT team's methods, it seems similar to what I did after a quick glance)
# doublecheck everything with data

# 7- ENERGY
# Energy supply data for Québec, a few things are missing/unclear, review the whole sheet, use raw data that is in the QCIOT data/inputs/energy SUPPLY data for QC/

# Figure the heat situation
# Make sure you got the right energy intensities (the economic output term I am uncertain of it)
# Correct energy intensities for electricity (remove energy of electricity produced by fossil fuels (see Eneko's email))

# SEE WORKPLAN_ETIENNE FOR ADDITIONNAL
# WHICH SECTORS are important for Québec?
# (***NA values when writing to excel, make sure to eliminate these for the input files to MEDEAS)



# 10- THINGS THAT REQUIRE ASKING MONICA/JORDI
# -GVT expenditures: (GE) are scarce for QC the way I calculated it, much more than for EUROPE/CATALONIA;
# ANSWER: They use EUREGIO weights for GE. (col EZS of EURegionallOtable_2010 - these are not weights, they are $ values. Look up how EUREGIO got these numbers)
#         It seems that it's a different accounting methodology. Catalonia gvt invests 10% of total GE in other sectors, 90% in 'Non Market services' (which is all that is public). Wonder what this 10% is, subsidies?
#         Could use Canada weights from EUREGIO and reallocate Quebec GE across sectors according to these weights.

# -Where did Monica/Francisca get their World FD category data? (World GFCF, GE, HH, INV (as/csGFCFwArray, as/csHHwArray, as/csGEwArray, as/csINVwArray))
# because my totals are not the same as theirs

# CHECKED Related question: is the GDP deflator (note tab) something to apply to economic data (other tabs)?
# ANSWER: The values are already deflated to 1995US.

# CHECKED Also: in WIOT, there is CONSumption by non-profits, where would that fit? (GFCF, HH, GE, INV?) I suspect HH or GE.
# ANSWER: MEDEAS-CAT includes non-profits expenses in HH.

# CHECKED -Historical GDP in economy.xlsx
# What is it calculated from?
# WORLD tab: perfectly adds up to GFCF + HH + GE + INV
# EUROPE tab: neither GFCF + HH + GE + INV + Exports FINAL GOODS nor LAB + CAP
# maybe it's VA (GDP as calculated from income = sum of Value Added across Sectors = sum of Capital Compensation + Labour Compensation + Net Tax across Sectors) as per MEDEAS D4.2 p.53
# ANSWER: GDP = sum of VA = CAP + LAB + Net tax + international trade margin
# =============================================================================
