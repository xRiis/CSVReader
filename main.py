import glob, os
import numpy as np
import pandas as pd

find_average = True
find_sum = True
find_region_area = True
find_area_proportion = True
skipped_rows = 11
column = 7
find_val_offset = 4
find_val_column = 1
fileDict = {}
dataDict = {}
fileList = []


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
        newList = appendBlank(newList, "Avg. Area", calcAverage(x))
        newLength -= 3
    if find_sum == True:
        newList = appendBlank(newList, "Sum Area", calcSum(x))
        newLength -= 3
    if find_region_area == True:
        regionArea = findVal(path, (length + find_val_offset), find_val_column, x)
        newList = appendBlank(newList, "Reg. Area", regionArea)
        newLength -= 3
        if find_area_proportion == True:
            sum = calcSum(x)
            regionArea = float(regionArea)
            quot = sum / regionArea
            newList = appendBlank(newList, "Area Prop.", quot)
            newLength -= 3
    for i in range(newLength):
        newList = np.append(newList, "")
    return newList


def writeToFile(path):
    writer = pd.ExcelWriter('file.xlsx', engine='xlsxwriter')
    for i in range(len(fileList)):
        i_len = len(fileDict[fileList[i]])
        total_data = generateData(i, i_len, path)
        df = pd.DataFrame({"Area": fileDict[fileList[i]],
                           "Data": total_data})
        df.to_excel(writer, sheet_name=fileList[i])
        dataDict[fileList[i]] = df
    writer.save()


if __name__ == '__main__':
    directory = str(input("Please enter directory containing all CSV files to be read: "))
    findFiles(directory)
    writeToFile(directory)
