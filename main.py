import glob, os
import numpy as np
import pandas as pd


global value_count

find_area_proportion, find_average, find_region_area, find_sum, write_summary = (True,) * 5

column = 7
skipped_rows = 11

find_val_column = 1
find_val_offset = 4

dataDict = {}
fileDict = {}
valueDict = {}

averageList, fileList, proportionList, regionAreaList, sumList = ([] for i in range(5))


def appendBlank(list, title, val):
    list = np.append(list, title)
    list = np.append(list, val)
    list = np.append(list, "")
    return list


def calcAverage(x):
    avg = fileDict[fileList[x]].mean()
    return avg


def calcSum(x):
    tot = fileDict[fileList[x]].sum()
    return tot


def findVal(path, row, col, x):
    os.chdir(path)
    tempDict = {}
    tempList = []
    for file in glob.glob("*.csv"):
        vals = pd.read_csv(file, skiprows=skipped_rows)
        tempDict[file] = vals.iloc[:, col].values
        tempList.append(file)
    val = tempDict[tempList[x]][row]
    return val


def findFiles(path):
    os.chdir(path)
    for file in glob.glob("*.csv"):
        fileList.append(file)
        newData = pd.read_csv(file, skiprows=skipped_rows)
        fileDict[file] = newData.iloc[:, column].values
        fileDict[file] = fileDict[file][~np.isnan(fileDict[file])]


def generateData(x, length, path):
    newList = []
    newLength = length
    if find_average == True:
        avg = calcAverage(x)
        newList = appendBlank(newList, "Avg. Area", avg)
        averageList.append(avg)
        newLength -= 3
    if find_sum == True:
        sum = calcSum(x)
        newList = appendBlank(newList, "Sum Area", sum)
        newLength -= 3
        sumList.append(sum)
    if find_region_area == True:
        region_area = findVal(path, (length + find_val_offset), find_val_column, x)
        regionArea = float(region_area)
        newList = appendBlank(newList, "Reg. Area", region_area)
        newLength -= 3
        regionAreaList.append(region_area)
        if find_area_proportion == True:
            sum = calcSum(x)
            quot = sum / regionArea
            newList = appendBlank(newList, "Area Prop.", quot)
            newLength -= 3
            proportionList.append(quot)
    for i in range(newLength):
        newList = np.append(newList, "")
    dataDict["Avg. Area"] = averageList
    dataDict["Sum Area"] = sumList
    dataDict["Reg. Area"] = regionAreaList
    dataDict["Area Prop."] = proportionList
    return newList


def writeSummary(engine):
    df = pd.DataFrame(dataDict)
    df.to_excel(engine, sheet_name="Summary")


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
    value_count = 0
    directory = str(input("Please enter directory containing all CSV files to be read: "))
    findFiles(directory)
    writeToFile(directory)
