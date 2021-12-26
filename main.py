import sys
import os
import cv2
import pytesseract
import openpyxl
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication

# вызов GUI
Form, Window = uic.loadUiType('pyqt.ui')

app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

def base():

    # Блок камеры
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        cv2.imshow('Video', frame)
        if cv2.waitKey(1) & 0xFF == ord('1'):
            cv2.imwrite('cam.png', frame)
            break
    cap.release()
    cv2.destroyAllWindows()

    # Вызов изображения
    img = cv2.imread('cam.png')

    # Блок работы с цветом изображения
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Считывание текста с картинки  ORC
    result = []
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-ORC\\tesseract'
    result = pytesseract.image_to_string(gray, lang='rus+en', config='-c page_separator=" "')
    ras = []
    ras = result.split()

    # Считывание Excel
    wb = openpyxl.reader.excel.load_workbook(filename='BAZA.xlsx')
    wb.active = 0
    sheet = wb.active
    NUMBase = []
    for i in range(1, 100):
        NUMBase += [(sheet['A' + str(i)].value)]

    # Проверка Базы
    ans = ''
    l = len(ras)
    if l < 1:
        ans = 'Номер не найден!'
    else:
        for i in range(l):
            if ras[i] in NUMBase:
                ans = 'Въезд разрешен!'
            else:
                ans = 'Въезд запрещён!'

    form.label_6.setText(ans)

def Excel():

    # открытие Excel
    a = os.system('BAZA.xlsx')
    return a

form.Start.clicked.connect(base) # Кнопка Старт
form.Excel.clicked.connect(Excel) # Кнопка Excel
form.Exit.clicked.connect(sys.exit) # Кнопка Выход

app.exec()

