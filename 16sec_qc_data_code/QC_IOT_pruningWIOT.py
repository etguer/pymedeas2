# This script reads WIOT Input-Output Tables, extracts the Canadian and Rest of World rows/columns
# and writes pruned files

# import packages
import pandas as pd
import numpy as np
import os.path

# Change working directory (Étienne MacOS)
os.chdir('/Users/Etguer/pymedeasQCdata/QC_IOT_data/QC_IOT_pruningWIOTpy_input')  # need to use '/' or '//' instead of '\'
# Change working directory (Étienne laptop)
# os.chdir('C:/Users/user/PycharmProjects/MEDEAS_QC_data')
wd = os.getcwd()

# set years
# this is useful for the getyearIndex function to always use 2004-2018 inclusively
all_years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])

# define data structures used further
rawdata = pd.DataFrame(data=None)
rawdata_list = list()

# set nrows/cols for 2004-2008 & 2009-2017 for intermediate consumption (IC)
rcOldIC1 = 0
rcOldIC2 = 25
rcNewIC1 = 0
rcNewIC2 = 32

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

def getrawdata(dr, yr, data_name):  # dr = directory; yr = year; data_name = PBM/intérieur if empty, otherwise can be
    # ImportationsTotales, ...
    stringname0 = 'WIOT_ROW/'
    stringname1 = 'ES symetriques provinciaux S QC '
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
        print('stringname:', stringname, 'sheet_name:', sheet_name, 'skiprows:', row_skip, 'usecols:', all_cols,
              'nrows:', nrows)
        tempdata = pd.read_excel(stringname, sheet_name=sheet_name, index_col=None,
                                 header=None, skiprows=row_skip, usecols=all_cols, nrows=nrows)
        tempdata = tempdata.reset_index(drop=True)  # reset row index
        tempdata.columns = range(tempdata.shape[1])  # reset column index
        return tempdata
    else:
        print(data_name, 'does not exist in', stringname)
        return

# get yearIndex (2004 = 0, 2018 = 13)
def getyearindex(yrs, yr):
    output = np.where(yrs == yr)[0][0]  # index of year in years
    return output

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

# get raw data
rawdataTP_World_list = list()  # total production world (for A matrix for QC IE)
rawdataE_Canada_list = list()  # final and intermediate exports (Canada - RoW, used for proportions for Quebec)
rawdataFD_World_list = list() # Final demand (HH+GFCF+GE+INV) for World


years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])
for year in years:
# for year in np.array([2004]):
    setyeartype(year)
    print('yeartype:', yeartype)
    print('yearIndex:', getyearindex(years, year))

    if year == 2009:
        rawdataFD_World = pd.DataFrame(data=None, index=range(2472), columns=range(220))
        rawdataFD_World_list.append(rawdataFD_World)
        rawdataE_Canada = pd.DataFrame(data=None, index=range(56), columns=range(2685))
        rawdataE_Canada_list.append(rawdataE_Canada)
    else:
        if year != 2009:
            rawdataFD_World = getrawdata(wd, year, data_name='FD_World')
            rawdataFD_World_list.append(rawdataFD_World)
            rawdataE_Canada = getrawdata(wd, year, data_name='Canada_Exports')
            rawdataE_Canada_list.append(rawdataE_Canada)



years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])
# array for already intérieur data for World (2004-2014)
asICworldArray = np.full(shape=(56, 56, 11), fill_value=np.nan, dtype=float)
# array for all sectors total production WORLD (2004-2014)
asTPworldArray = np.full(shape=(56, 1, 11), fill_value=-1, dtype=float)
# array for all sectors world value added
asVAworldArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)

# array for all sectors Gross Fixed Capital Formation (added to single column) WORLD
asGFCFwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors Household demand (added to single column) WORLD
asHHwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors Government Expenditure (added to single column) WORLD
asGEwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)
# array for all sectors Inventories (added to single column) WORLD
asINVwArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)

# array for all canada sectors intermediate exports and final demand
# (CAN has 56 sectors in IE, 5 categories in final demand)
asIEFEcanArray = np.full(shape=(56, 2685 - 56 - 5, len(years)), fill_value=0, dtype=float)

# array for all sectors World minus QC (initially this is World and then QC is subtracted) Final Demand
asFDrowArray = np.full(shape=(56, len(years)), fill_value=0, dtype=float)

years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018])

# get the
for year in years:
# for year in np.array([2004]):
    setyeartype(year)
    df_to_array(df_list=rawdataFD_World_list, data_name='FDw', yr=year)  # creates asFDrowArray CHECK
    df_to_array(df_list=rawdataE_Canada_list, data_name='IE_CAN', yr=year)  # creates asIEFEcanArray CHECK



# FILL GAPS (2009 and 2012-2013) - time linearity assumption
# 2009
asIEFEcanArray[:,:,5] = (asIEFEcanArray[:,:,4]+asIEFEcanArray[:,:,6])/2
asFDrowArray[:,5] = (asFDrowArray[:,4]+asFDrowArray[:,6])/2

#--------------------------------------------------------------------------------
# read World IC raw data from 2004-2014 (2009 skipped with nans)
rawdataICworld_list = list()  # intermediate consumption (World to World industry exchanges)
row_init = 7
year_one_w = 2004
# for year in np.array([2004]):
for year in range(2004, 2015):
    print('Reading raw data (IC world) for year', year)
    if year == 2009:
        rawdataICworld = pd.DataFrame(data=None, index=range(2464), columns=range(2464))
        rawdataICworld_list.append(rawdataICworld)
    else:
        yearIndex = year - year_one_w
        skiprows = row_init - 1
        stringname = 'WIOT_ROW/WIOT'+str(year)+'_Nov16_ROW.xlsb'
        rawdataICworld = pd.read_excel(stringname, sheet_name=str(year), index_col=None,
                                   header=None, skiprows=skiprows, usecols='E:CPX', nrows=2464)
        rawdataICworld = rawdataICworld.reset_index(drop=True)  # reset row index
        rawdataICworld.columns = range(rawdataICworld.shape[1])  # reset column index
        rawdataICworld_list.append(rawdataICworld)

# --------------------------------------------------------------------------------
# aggregate IC array across countries
# for year in np.array([2004]):
for year in range(2004, 2015):
    yearIndex = year - 2004
    print('Aggregating raw data (IC world) across countries for year', year)
    if year == 2009:
        pass
    else:
        rawdata = rawdataICworld_list[yearIndex]
        rawdatar = rawdata.to_numpy().reshape(len(rawdata.index), len(rawdata.columns))
        datar = np.full(shape=(56, 56), fill_value=0, dtype=float)
        for i in range(0, 43+1):
            for j in range(0, 43+1):
                tempr = rawdatar[(i * 56):(55 + 1 + (56 * i)), (j * 56):(55 + 1 + (56 * j))]
                datar = datar + tempr
        asICworldArray[:, :, yearIndex] = datar

#Fill GAPS (2009) (Time linearity assumption)
asICworldArray[:,:,5] = (asICworldArray[:,:,4]+asICworldArray[:,:,6])/2


# read World TP raw data from 2004-2014 (2009 skipped with nans)
rawdataTPworld_list = list()  # intermediate consumption (World to World industry exchanges)
row_init = 7
year_one_w = 2004
# for year in np.array([2004]):
for year in range(2004, 2015):
    print('Reading raw data (TP world) for year', year)
    if year == 2009:
        rawdataTPworld = pd.DataFrame(data=None, index=range(2464), columns=range(1))
        rawdataTPworld_list.append(rawdataTPworld)
    else:
        yearIndex = year - year_one_w
        skiprows = row_init - 1
        stringname = 'WIOT_ROW/WIOT'+str(year)+'_Nov16_ROW.xlsb'
        rawdataTPworld = pd.read_excel(stringname, sheet_name=str(year), index_col=None,
                                   header=None, skiprows=skiprows, usecols='CYK', nrows=2464)
        rawdataTPworld = rawdataTPworld.reset_index(drop=True)  # reset row index
        rawdataTPworld.columns = range(rawdataTPworld.shape[1])  # reset column index
        rawdataTPworld_list.append(rawdataTPworld)

# aggregate TP array across countries
# for year in np.array([2004]):
for year in range(2004, 2015):
    yearIndex = year - 2004
    print('Aggregating raw data (TP world) across countries for year', year)
    if year == 2009:
        pass
    else:
        rawdata = rawdataTPworld_list[yearIndex]
        rawdatar = rawdata.to_numpy().reshape(len(rawdata.index), len(rawdata.columns))
        datar = np.full(shape=(56, 1), fill_value=0, dtype=float)
        for i in range(0, 43+1):
            tempr = rawdatar[(i * 56):(55 + 1 + (56 * i))]
            datar = datar + tempr
        asTPworldArray[:, :, yearIndex] = datar

# FILL GAP (2009) (time linearity assumption)
asTPworldArray[:,:,5] = (asTPworldArray[:,:,4]+asTPworldArray[:,:,6])/2



# ------GET World VA data from 2004-2014 (2009 skipped)------
year_one_w = 2004
# for year in np.array([2004]):
for year in range(2004, 2014+1):
    print('Reading raw data (VA world) for year', year)
    if year == 2009:
        pass
    else:
        yearIndex = year - year_one_w
        stringname = 'WIOT_ROW/WIOT'+str(year)+'_Nov16_ROW.xlsb'
        rawdataVAworld = pd.read_excel(stringname, sheet_name=str(year), index_col=None, header=list([2, 3, 4, 5]))
        rawdata = rawdataVAworld.iloc[2469, 4:2465]  # take the VA row and columns
        rawdatar = rawdata.groupby(level=[0]).sum()  # aggregate across countries, retain sectors and CAP/LAB/VA
        tempar = rawdatar.to_numpy().reshape(len(rawdatar.index))
        asVAworldArray[:, yearIndex] = rawdatar

# FILL GAP (2009) (time linearity assumption)
asVAworldArray[:,5] = (asVAworldArray[:,4]+asVAworldArray[:,6])/2



# ------GET World GFCF, GE, HH, INV (aswArray, aswArray, aswArray, aswArray)  data from 2004-2014 (2009 skipped)------
year_one_w = 2004
year_last = 2014
skiprows = 0
index_col = list([0, 1, 2, 3])

# make exclusion list for non FD columns (4th level (so level=3) of headers, from c1 to c56)
exclude_list = list()
for x in range(1, 56+1):
    label = 'c'+str(x)
    exclude_list.append(label)

exclude_list.append('c62')  # also exclude total output (last column)

# for year in np.array([2004]):
for year in range(year_one_w, year_last+1):
    print('Reading raw data (FD categories world) for year', year)
    if year == 2009:
        pass
    else:
        yearIndex = year - year_one_w
        stringname = 'WIOT_ROW/WIOT'+str(year)+'_Nov16_ROW.xlsb'
        rawdataFDworld = pd.read_excel(stringname, sheet_name=str(year), index_col=index_col,
                                       header=list([2, 3, 4, 5]))
        rawdata = rawdataFDworld.iloc[:-8, :]  # take the FD rows (exclude the last 8 'total' rows)
        rawdata1 = rawdata.drop(labels=exclude_list, axis=1, level=3) # exclude up to c56 (FD starts at c57 for level 4 header)
        rawdatar = rawdata1.groupby(axis=0, level=[0]).sum()  # aggregate countries together (input/row side), retain sectors
        rawdatar1 = rawdatar.groupby(axis=1, level=[0]).sum()  # aggregate countries together (output/column side), retain final demand categories
        # *******CONS_np, consumption of non-profits, not sure which category to add it to, gotta check with WORLD values (with different sectorilazition)*******
        asGEwArray[:, yearIndex] = rawdatar1['CONS_g'].to_numpy().reshape(len(rawdatar1.index))
        asHHwArray[:, yearIndex] = rawdatar1['CONS_h'].to_numpy().reshape(len(rawdatar1.index)) + rawdatar1['CONS_np'].to_numpy().reshape(len(rawdatar1.index))
        asGFCFwArray[:, yearIndex] = rawdatar1['GFCF'].to_numpy().reshape(len(rawdatar1.index))
        asINVwArray[:, yearIndex] = rawdatar1['INVEN'].to_numpy().reshape(len(rawdatar1.index))

# %%%%%%%%%%%%%%%%%%%%%

# Fill 2009 GAP (time linearity assumption)
asGFCFwArray[:,5] = (asGFCFwArray[:,4]+asGFCFwArray[:,6])/2
asHHwArray[:,5] = (asHHwArray[:,4]+asHHwArray[:,6])/2
asGEwArray[:,5] = (asGEwArray[:,4]+asGEwArray[:,6])/2
asINVwArray[:,5] = (asINVwArray[:,4]+asINVwArray[:,6])/2


#WRITE these dataframes to an excel workbook 'pruningTest.xlsx'

years = np.array([2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014])
#path = '/Users/Etguer/Documents/QC_IOT_data/input/pruningTest.xlsx'

# for year in np.array([2004]):
for year in years:
    path = '/Users/Etguer/pymedeasQCdata/QC_IOT_data/QC_IOTpy_input/prunedData'+str(year)+'.xlsx'
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        yearIndex = year - 2004
        data = pd.DataFrame(data=asICworldArray[:, :, yearIndex])
        data.to_excel(writer, sheet_name='asICworldArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
        data = pd.DataFrame(data=asTPworldArray[:, :, yearIndex])
        data.to_excel(writer, sheet_name='asTPworldArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
        data = pd.DataFrame(data=asIEFEcanArray[:, :, yearIndex])
        data.to_excel(writer, sheet_name='asIEFEcanArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)

path = '/Users/Etguer/pymedeasQCdata/QC_IOT_data/QC_IOTpy_input/prunedDataOther.xlsx'
with pd.ExcelWriter(path, engine='openpyxl') as writer:
    data = pd.DataFrame(data=asVAworldArray[:, :])
    data.to_excel(writer, sheet_name='asVAworldArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
    data = pd.DataFrame(data=asGFCFwArray[:, :])
    data.to_excel(writer, sheet_name='asGFCFwArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
    data = pd.DataFrame(data=asHHwArray[:, :])
    data.to_excel(writer, sheet_name='asHHwArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
    data = pd.DataFrame(data=asGEwArray[:, :])
    data.to_excel(writer, sheet_name='asGEwArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
    data = pd.DataFrame(data=asINVwArray[:, :])
    data.to_excel(writer, sheet_name='asINVwArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
    data = pd.DataFrame(data=asFDrowArray[:, :])
    data.to_excel(writer, sheet_name='asFDrowArray', na_rep='na', header=False, index=False, startrow=0, startcol=0)
