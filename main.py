import glob, os, statistics
import numpy as np
import pandas as pd


# Do we want to do all these things? Yes
find_area_proportion, find_average, find_region_area, find_standard_deviation, find_sum, write_summary = (True,) * 6

column = 7 # Column to read from, begins at 0
skipped_rows = 11 # Number of rows to skip so the reader doesn't get confused, begins at 0

# Where on the CSV do we find the value for region area? Not in the lists
find_val_column = 1 # Second column from the left
find_val_offset = 4 # Four rows after we stop reading

# Create a bunch of empty dicts to be used later
dataDict = {}
fileDict = {}
valueDict = {}

# Create a bunch of empty lists to be used later
averageList, fileList, proportionList, regionAreaList, sdvList, sumList = ([] for i in range(6))


# Generate blank spaces so lists in our dicts don't have mismatched dimensions
def appendBlank(list, title, val):
    list = np.append(list, title)
    list = np.append(list, val)
    list = np.append(list, "")
    return list


# Find the average of the column we're reading from inside one of our files
def calcAverage(x):
    avg = fileDict[fileList[x]].mean()
    return avg


# Find the standard deviation of the column we're reading from inside one of our files
def calcStDev(x):
    sdv = statistics.stdev(fileDict[fileList[x]])
    return sdv


# Find the sum of the column we're reading from inside one of our files
def calcSum(x):
    tot = fileDict[fileList[x]].sum()
    return tot


# Find the value of a specific cell
def findVal(path, row, col, x):
    os.chdir(path)
    tempDict = {}
    tempList = []
    for file in glob.glob("*.csv"):
        vals = pd.read_csv(file, skiprows=skipped_rows)
        tempDict[file] = vals.iloc[:, col].values # Scan in the whole column
        tempList.append(file)
    val = tempDict[tempList[x]][row] # Pick the value we want out of that column. Not efficient but it works
    return val


# Fill a dict with names of files and their data
def findFiles(path):
    os.chdir(path)
    for file in glob.glob("*.csv"):
        fileList.append(file)
        newData = pd.read_csv(file, skiprows=skipped_rows)
        fileDict[file] = newData.iloc[:, column].values
        fileDict[file] = fileDict[file][~np.isnan(fileDict[file])]


# Fill a dictionary with whatever data we want to find
def generateData(x, length, path):
    newList = [] # Temp list
    newLength = length
    if find_average == True:
        avg = calcAverage(x)
        newList = appendBlank(newList, "Avg. Area", avg)
        averageList.append(avg)
        newLength -= 3 # We wrote three lines, so subtract that amount from the list to avoid a dimensional mismatch
        dataDict["Avg. Area"] = averageList
    if find_sum == True:
        sum = calcSum(x)
        newList = appendBlank(newList, "Sum Area", sum)
        newLength -= 3
        sumList.append(sum)
        dataDict["Sum Area"] = sumList
    if find_region_area == True:
        region_area = findVal(path, (length + find_val_offset), find_val_column, x)
        regionArea = float(region_area)
        newList = appendBlank(newList, "Reg. Area", region_area)
        newLength -= 3
        regionAreaList.append(region_area)
        dataDict["Reg. Area"] = regionAreaList
        if find_area_proportion == True: # We won't care about this if we don't care about the total region area
            sum = calcSum(x)
            quot = sum / regionArea
            newList = appendBlank(newList, "Area Prop.", quot)
            newLength -= 3
            proportionList.append(quot)
            dataDict["Area Prop."] = proportionList
    if find_standard_deviation == True:
        stdev = calcStDev(x)
        newList = appendBlank(newList, "St. Dev.", stdev)
        newLength -= 3
        sdvList.append(stdev)
        dataDict["St. Dev."] = sdvList
    for i in range(newLength):
        newList = np.append(newList, "")
    return newList


# Creates an individual summary sheet
def writeSummary(engine):
    df = pd.DataFrame(dataDict)
    df.to_excel(engine, sheet_name="Summary")


# Iterate through all of our files, trawl through their data, and write them to individual worksheets
def writeToFile(path):
    writer = pd.ExcelWriter("file.xlsx", engine="xlsxwriter")
    for i in range(len(fileList)):
        i_len = len(fileDict[fileList[i]])
        total_data = generateData(i, i_len, path)
        if write_summary == True:
            writeSummary(writer)
        df = pd.DataFrame({"Area": fileDict[fileList[i]],
                           "Data": total_data})
        df.to_excel(writer, sheet_name=fileList[i])
    writer.save()


if __name__ == '__main__':
    directory = str(input("Please enter directory containing all CSV files to be read: "))
    findFiles(directory)
    writeToFile(directory)
