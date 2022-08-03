# used to create parameter names in SimaPro

import openpyxl as oxl
import pyautogui
import time

pyautogui.PAUSE = 2
pyautogui.FAILSAFE = True


def readPars(filename,directory,targetsheet):
    parlist = []
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
    column = sheet['B']
    for cell in column:
        parlist.append(cell.value)
    parlist.pop(0)
    return parlist

def keystrokes(par):
    pyautogui.hotkey('ctrl','i')
    pyautogui.write(par,interval=0.01)
    #pyautogui.press('enter')
    return

def loopPars(parlist):
    for i in parlist:
        keystrokes(i)
    return

def main():
    # add directory here
    directory = r""
    filename = 'Foodstep välipalat.xlsx'
    targetsheet = 'test'
    parlist = readPars(filename,directory,targetsheet)
    print(f'{len(parlist)} parameters in list.')
    tic = time.perf_counter()
    #print(parlist)
    loopPars(parlist)
    toc = time.perf_counter()
    print(f'Simapro input executed in {toc-tic:0.4f} seconds.')
    return


if __name__ == '__main__':
    main()