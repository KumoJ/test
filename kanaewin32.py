import cv2

#元画像の読み込み
img = cv2.imread('data/kanae.png')

img2 = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)

img3 = cv2.resize(img2, dsize=(54,96))

#画像の縦の画素数、横の画素数を取得
h,w = img3.shape

#３値化に使う低い側の閾値を決める
th_low_value=int(input('1つ目の閾値を入力してください。:'))
#３値化に使う高い側の閾値を決める
th_high_value=int(input('2つ目の数字を入力してください。:'))
#3値化の白黒以外の色
gray=100

import win32com.client as win32
from pathlib import Path
import os

#２次元配置の並べ替えとエクセルへの書き込み
    
def write_list_2d(sheet, list_2d, start_row, start_col):
    for y, row in enumerate(list_2d):
        for x, cell in enumerate(row):
            sheet = new_func()   
            sheet.Cells(start_row + y, start_col + x).Value = list_2d[y][x]

def new_func():
    sheet = wb.Worksheets(1)
    return sheet


# 既存のExcelファイルを開く

xl = win32.Dispatch('Excel.Application')
fpath = os.path.join(os.getcwd(),'雛形.xlsm')
wb = xl.Workbooks.Open(fpath)

wb = xl.ActiveWorkbook

# ドットの座標
dots_gray = []

for i in range(h):
    for j in range(w):
        if th_low_value <= img3[i,j] <=th_high_value:
            img3[i,j]=gray
            dots_gray.append((i,j))

# EXCELに書き込み
write_list_2d('Sheet1', dots_gray, 3, 12)


# ドットの座標
dots_black = []

for i in range(h):
    for j in range(w):
        if  img3[i,j] < th_low_value:
            img3[i,j]=0
            dots_black.append((i,j))


# EXCELに書き込み
write_list_2d('Sheet1', dots_black, 3,18 )

# 保存
new_filename = r'C:\Users\西村篤哉\デスクトップ\NC\outputA.xlsm'
wb.SaveAs(new_filename, FileFormat=52)  # 52はxlsm形式

A = max(len(dots_gray),len(dots_black))
list_odds = []
list_even = []
list_coron = []
list_N = []
list_borderX = []
list_borderY = []
list_capX = []
list_capY = []
list_pitchX = []
list_pitchY = []
list_NCg = ['O0','G00','G90']
list_NCo = ['O1','M98','P1','L1']
list_NCg_new = []
list_NCo_new = []

bX = input('長手のボーダーを入力してください:')
bY = input('短手のボーダーを入力してください:')

pX = input('長手のピッチを入力してください:')
pY = input('短手のピッチを入力してください:')


for odds in range(3,int(A+2)*2,2):
    list_odds.append(odds)

for even in range(4,int(A+2)*2,2):
    list_even.append(even)

for coron in range (A):
    coron = ':'
    list_coron.append(coron)
    

for N in range(A):
    N = int(900)
    list_N.append(N)

for borderX in range(A):
    borderX = int(bX)
    list_borderX.append(borderX)

for borderY in range(A):
    borderY = int(bY)
    list_borderY.append(borderY)

for capX in range(A):
    capX = 'X'
    list_capX.append(capX)

for capY in range(A):
    capY = 'Y'
    list_capY.append(capY)

for pitchX in range(A):
    pitchX = int(pX)
    list_pitchX.append(pitchX)

for pitchY in range(A):
    pitchY = int(pY)
    list_pitchY.append(pitchY)

for _ in range(A):
    list_NCg_new.append(list_NCg)

for _ in range(A):
    list_NCg_new.append(list_NCo)

wb = xl.ActiveWorkbook
sheet = wb.Worksheets(1)

for i in range(0,A):
    sheet.Cells(row=i+3,column=32,value=list_capX[i])

    sheet.Cells(row=i+3,column=33,value=list_capY[i])

    sheet.Cells(row=i+3,column=24,value=list_odds[i])

    sheet.Cells(row=i+int(A)+3,column=24,value=list_even[i])

    sheet.Cells(row=i+3,column=25,value=list_coron[i])

    sheet.Cells(row=i+int(A)+3,column=25,value=list_coron[i])

    sheet.Cells(row=i+3,column=5,value=list_borderX[i])

    sheet.Cells(row=i+3,column=6,value=list_pitchX[i])

    sheet.Cells(row=i+3,column=8,value=list_N[i])

    sheet.Cells(row=i+3,column=9,value=list_borderY[i])

    sheet.Cells(row=i+3,column=10,value=list_pitchY[i])

for row_index, row_values in enumerate(list_NCg_new):
    for col_index, value in enumerate(row_values):
        cell = sheet.cells(row = row_index+3, column = col_index+27)
        cell.value = value

for row_index, row_values in enumerate(list_NCo_new):
    for col_index, value in enumerate(row_values):
        cell = sheet.Cells(row = row_index+int(A)+3, column = col_index+27)
        cell.value = value

wb.SaveAs(new_filename, FileFormat=52)  # 52はxlsm形式

xl.Quit()
