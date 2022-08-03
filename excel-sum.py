# script for reading a specific Excel-sheet containing survey responses from a specific folder
import os
import openpyxl as oxl
import pandas as pd
from collections import defaultdict
import time

def listFiles(directory):
    filelist = os.listdir(directory)
    print(f'{len(filelist)} file(s) in folder.')
    return filelist

def readExcel(filename,directory,targetsheet):
    try:
        wb = oxl.load_workbook(directory+'\\'+filename, data_only=True)
    except oxl.utils.exceptions.InvalidFileException:
        print(f'{filename} is not a valid file.')
        return
    try:
        sheet = wb[targetsheet]
    except KeyError:
        print(f'{filename} does not contain "{targetsheet}" sheet.')
        return
    dic = {col[0].value: [cell.value for cell in col[1:]] for col in sheet.columns}
    return dic

def loopFiles(filelist,directory,targetsheet):
    templist = []
    dic = defaultdict(list)
    for file in filelist:
        singledic = readExcel(file,directory,targetsheet)
        templist.append(singledic)
    for i in templist:
        try:
            for k,v in i.items():
                dic[k].extend(v)
        except AttributeError:
            continue
    return dic

def toDf(dic):
    df = pd.DataFrame.from_dict(dic,orient='index').transpose()
    df = df[df['Lohkotunnus_lohkotunnus'] != 0].reset_index(drop=True)
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.applymap(lambda x: x.capitalize() if type(x) == str else x)
    #print(df.applymap(type).eq(str).all())
    df.to_csv('Hiililounas-farmer-data.csv',index=False)
    return 

def main():
    # add directory here
    directory = r""
    targetsheet = 'Kooste'
    tic = time.perf_counter()
    filelist = listFiles(directory)
    dic = loopFiles(filelist,directory,targetsheet)
    toDf(dic)
    toc = time.perf_counter()
    print(f'Executed in {toc-tic:0.4f} seconds.')
    return


if __name__ == '__main__':
    main()