import pyautogui
import time
import xlrd
import pyperclip

# Mouse click event
# pyautogui : https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes, leftOrRight, img, reTry):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence = 0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks = clickTimes, interval = 0.2, duration = 0.2, button = leftOrRight)
                break
            print("Can't find the IMAGE, will retry after 0.1 second")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence = 0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks = clickTimes, interval = 0.2, duration = 0.2, button = leftOrRight)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence = 0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks = clickTimes, interval = 0.2, duration = 0.2, button = leftOrRight)
                print("Repeat.")
                i  += 1
            time.sleep(0.1)


# Data Reader
# cmdType.value  1 left-click
#                2 left-double-click
#                3 right-lick
#                4 input
#                5 wait
#                6 scroll
# ctype     empty   0
#           string  1
#           number  2
#           date    3
#           bool    4
#           error   5
def dataCheck(sheet1):
    checkCmd = True
    #check row count
    if sheet1.nrows < 2:
        print("This sheet is empty")
        checkCmd = False
    # Check every row
    i = 1
    cmdTypeList = [1, 2, 3, 4, 5, 6]
    while i < sheet1.nrows:
        # The first column
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or cmdType.value not in cmdTypeList:
            print('The ', i + 1, "st row, 1st column's data got something wrong")
            checkCmd = False
        # The second column
        cmdValue = sheet1.row(i)[1]
        # Reading-image-type & Clicking-type must be a string
        if cmdType.value == 1 or cmdType.value == 2 or cmdType.value == 3:
            if cmdValue.ctype != 1:
                print('The ', i + 1, "ed row, 2ed column's data got something wrong")
                checkCmd = False
        # The input-type can't be empty
        if cmdType.value == 4:
            if cmdValue.ctype == 0:
                print('The ', i + 1, "rd row, 2ed column's data got something wrong")
                checkCmd = False
        # Waiting-type must be a number
        if cmdType.value == 5:
            if cmdValue.ctype != 2:
                print('The ', i + 1, "th row, 2ed column's data got something wrong")
                checkCmd = False
        # Scrolling-type must be a number
        if cmdType.value == 6:
            if cmdValue.ctype != 2:
                print('The ', i + 1, "th row, 2ed column's data got something wrong")
                checkCmd = False
        i  += 1
    return checkCmd


# Get command's loop-time config
def getLoopTime(row, sheet):
    columnIndex = 2
    loop = 1 
    if sheet.row(row)[columnIndex].ctype == 2 and sheet.row(row)[columnIndex].value != 0 :
        loop = sheet.row(row)[columnIndex].value
    return loop


# Main
def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        # Get current row's operation config
        cmdType = sheet1.row(i)[0]

        # [1] refferrers to CLICK
        if cmdType.value == 1:
            # Get image name
            img = sheet1.row(i)[1].value
            loop = getLoopTime(i, sheet1)
            mouseClick(1, "left", img, loop)
            print("Click ", img)

        # [2] refferrers to DOUBLE CLICK
        elif cmdType.value == 2:
            # Get image name
            img = sheet1.row(i)[1].value
            loop = getLoopTime(i, sheet1)
            mouseClick(2, "left", img, loop)
            print("Double click ", img)

        # [3] refferrers to RIGHT CLICK
        elif cmdType.value == 3:
            # Get image name
            img = sheet1.row(i)[1].value
            loop = getLoopTime(i, sheet1)
            mouseClick(1, "right", img, loop)
            print("Right click ", img) 

        # [4] refferrers to INPUT
        elif cmdType.value == 4:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("Input : ", inputValue)

        # [5] refferrers to PAUSE
        elif cmdType.value == 5:
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("Pause ", waitTime, " seconds")

        # [6] refferrers to SCROLL
        elif cmdType.value == 6:
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("scroll ", int(scroll))

        i  += 1

if __name__ == '__main__':
    file = 'cmd.xls'
    # Open workbook
    workBook = xlrd.open_workbook(filename = file)
    # Get sheet data by using index
    sheet1 = workBook.sheet_by_index(0)
    # Check Data
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        key = input('What are you going to do: \n 1.Just ONCE \n 2.Until manually break \n')
        if key == '1':
            # Get every operation config
            mainWork(sheet1)
        elif key == '2':
            while True:
                mainWork(sheet1)
                time.sleep(0.1)
                print("0.1 second break")    
    else:
        print('Input error OR script exit')
