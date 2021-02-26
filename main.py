import glob, os
import numpy as np
import pandas as pd

skipped_rows = 11
column = 7
fileDict = {}
dataDict = {}
fileList = []


def appendBlank(x, col):
    i_len = len(fileDict[fileList[x]]) - 1
    for i in range(i_len):
        col = np.append(col, np.nan)
    return col


def calcAverage(x):
    avg = fileDict[fileList[x]].mean()
    appendBlank(x, avg)
    return avg


def findFiles(path):
    os.chdir(path)
    for file in glob.glob("*.csv"):
        fileList.append(file)
        newData = pd.read_csv(file, skiprows=skipped_rows)
        fileDict[file] = newData.iloc[:, column].values
        fileDict[file] = fileDict[file][~np.isnan(fileDict[file])]


def writeToFile():
    writer = pd.ExcelWriter('file.xlsx', engine='xlsxwriter')
    for i in range(len(fileList)):
        average_area = calcAverage(i)
        df = pd.DataFrame({"Area": fileDict[fileList[i]],
                           "Avg. Area": average_area})
        df.to_excel(writer, sheet_name=fileList[i])
        dataDict[fileList[i]] = df
    writer.save()


if __name__ == '__main__':
    directory = str(input("Please enter directory containing all CSV files to be read: "))
    findFiles(directory)
    writeToFile()
