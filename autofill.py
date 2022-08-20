#!/usr/bin/env python3
from PyQt5 import QtCore, QtWidgets, uic
import sys
import os
import docx
from datetime import datetime, timedelta
from threading import Thread
from time import sleep
from urllib.request import urlopen
import re

import sys
if sys.platform.startswith("win"):
    WIN = True
    WIN_css = 'font-size: 11pt;'
else:
    WIN = False
    WIN_css = ''

APP_PATH = os.path.abspath(os.path.dirname(sys.argv[0]))

# диапазон дат четвертей
list_date = ['01.09.2021', '31.10.2021', '08.11.2021', '28.12.2021', '10.01.2022', '20.03.2022', '28.03.2022', '22.05.2022']
list_days = []
all_days = 0
dict_days = {}
week_days = []
double_days = []
file_name = ''
count_hours = 0
table_number = -1
one = 2
column_with_days = 3
list_one = []

version = 'Версия 1.1 от 20.08.2022'

# загрузка окна
app = QtWidgets.QApplication([])
win = uic.loadUi('main.ui')
win_about = uic.loadUi('about.ui')
win_about.labelVersion.setText(version)
win_help = uic.loadUi('help.ui')
win_help.groupBoxLink.hide()
win_help.resize(500, 300)

win.resize(100, 100)
win.labelVersion.setText(version)

# получаем список всех виджетов и создаем переменные
for i in dir(win):
    try:
        exec('win.'+i+'.objectName()')
        exec('globals()[win.'+i+'.objectName()] = win.'+i)
    except:
        pass

def fixFontIfWIN():
    if WIN:
        font_size = 'QWidget {font-size: 11pt;}'
        win.setStyleSheet(font_size)
        win.pushButtonFill.setStyleSheet('QWidget {font-size: 18pt;}')
        for i in range(1, 5):
            globals()['label0'+str(i)].setStyleSheet('QWidget {font-size: 12pt;}')
        win_about.setStyleSheet(font_size)
        win_about.label.setStyleSheet('QWidget {font-size: 17pt;}')
        win_help.setStyleSheet(font_size)

def loadDateFromFile():
    if os.path.exists(os.path.join(APP_PATH, 'сохранённый диапазон дат.txt')):
        with open(os.path.join(APP_PATH, 'сохранённый диапазон дат.txt'), 'r') as f:
            i = 0
            for line in f:
                list_date[i] = line.strip()
                i += 1
    print('Loaded from file')

def loadDateToApp():
    for i in range(8):
        qdate = QtCore.QDate.fromString(list_date[i], 'dd.MM.yyyy')
        globals()['dateEdit'+str(i)].setDate(qdate)

def connectSignals():
    for i in range(8):
        globals()['dateEdit'+str(i)].dateChanged.connect(lambda: changedDateEdit())
        globals()['dateEdit'+str(i)].setStyleSheet('QWidget {%s}'%WIN_css)
    for i in range(6):
        globals()['checkBoxWeek'+str(i)].clicked.connect(lambda: readWeekDays())
        globals()['spinBoxWeek'+str(i)].valueChanged.connect(lambda: readWeekDays())
    checkBoxYear.clicked.connect(lambda: readWeekDays())
    pushButton.clicked.connect(lambda: openFiles())
    pushButtonFill.clicked.connect(lambda: fill())
    frame_2.setStyleSheet('QWidget {}')
    frame_3.setStyleSheet('QWidget {}')
    comboBoxColumns.currentIndexChanged.connect(setColumnWithDay)
    comboBoxTables.textActivated.connect(setTable)
    action_1.triggered.connect(lambda: win_help.show())
    action_2.triggered.connect(lambda: win_about.show())
    win_help.pushButtonChekNewVersion.clicked.connect(lambda: buttonCheckNewVersionClick())

class CheckVersionThread(QtCore.QThread):
    labelStatusChangeSignal = QtCore.pyqtSignal(str)
    labelLinkShowSignal = QtCore.pyqtSignal(bool)
    def  __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
    def run(self):
        self.labelStatusChangeSignal.emit('Проверка...')
        self.labelLinkShowSignal.emit(False)
        new_ver = checkNewVersion()
        if new_ver:
            if new_ver == -1:
                self.labelStatusChangeSignal.emit('Произошла ошибка')
            else:
                self.labelStatusChangeSignal.emit('Доступна новая версия '+new_ver+'!')
                self.labelLinkShowSignal.emit(True)
        else:
            self.labelStatusChangeSignal.emit('Вы используете актуальную версию')

def on_label_status_change(s):
    win_help.label_status.setText(s)

def on_label_link_show(show):
    if show:
        win_help.groupBoxLink.show()
    else:
        win_help.groupBoxLink.hide()

check_version_thread = CheckVersionThread()
check_version_thread.labelStatusChangeSignal.connect(on_label_status_change, QtCore.Qt.QueuedConnection)
check_version_thread.labelLinkShowSignal.connect(on_label_link_show, QtCore.Qt.QueuedConnection)

def checkNewVersion():
    try:
        url = 'https://sourceforge.net/projects/autofillktp/files/'
        html = urlopen(url, timeout=5).read().decode('utf-8')
        s = re.findall('latest.*title=(.*?)\.zip', html)
        new_ver = s[0].split()[-1]
        old_ver = version.split()[1]
        if int(float(new_ver)*100) > int(float(old_ver)*100):
            return(new_ver)
    except:
        return(-1)
    return(False)

def buttonCheckNewVersionClick():
    check_version_thread.start()

def setTable(text):
    global table_number
    table_number = int(text.split()[-1])-1
    read_hours_thread.start()
    

def setColumnWithDay(index):
    if index <= 0:
        return
    global column_with_days
    column_with_days = index
    pushButtonFill.setText('Заполнить')
    pushButtonFill.setStyleSheet('QWidget {font-size: 18pt;}')
    for i in range(6):
        globals()['checkBoxWeek'+str(i)].setChecked(False)
        globals()['spinBoxWeek'+str(i)].setValue(1)
    readWeekDays()


def changedDateEdit():
    global list_days, dict_days, all_days
    for i in range(0, 7):
        d1 = globals()['dateEdit'+str(i)].dateTime().toString('dd.MM.yyyy')
        d2 = globals()['dateEdit'+str(i+1)].dateTime().toString('dd.MM.yyyy')

        if not checkValidDatesArr(d1, d2):
            globals()['dateEdit'+str(i)].setStyleSheet('QWidget {background-color: rgb(255, 190, 191);%s}'%WIN_css)
            globals()['dateEdit'+str(i+1)].setStyleSheet('QWidget {background-color: rgb(255, 190, 191);%s}'%WIN_css)
            list_days = []
            dict_days = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
            all_days = 0
            loadWeekDays()
            readWeekDays(err=True)
            frame_2.setEnabled(False)
            return
        else:
            globals()['dateEdit'+str(i)].setStyleSheet('QWidget {%s}'%WIN_css)
            globals()['dateEdit'+str(i+1)].setStyleSheet('QWidget {%s}'%WIN_css)
            frame_2.setEnabled(True)
    saveDateToFile()
    makeListOfDays()
    loadWeekDays()
    readWeekDays()

def loadWeekDays():
    for i in range(6):
        globals()['weekday'+str(i)].setText(str(dict_days[i]))
    label_all_days.setText(str(all_days))

def saveDateToFile():
    global list_date
    for i in range(8):
        list_date[i] = globals()['dateEdit'+str(i)].dateTime().toString('dd.MM.yyyy')
    with open(os.path.join(APP_PATH, 'сохранённый диапазон дат.txt'), 'w') as f:
        f.write('\n'.join(list_date))
    print('Saved to file')

def makeListOfDays(weekday=[], selected = False, doubleday=[]):
    global list_days, all_days, dict_days
    if selected:
        list_days = []
    all_days = 0
    dict_days = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
    list_all_days=(0, 1, 2, 3, 4, 5)
    year = ''
    if checkBoxYear.isChecked():
        year = '.%y'
    for i in range(0, 7, 2):
        d1 = datetime.strptime(list_date[i], '%d.%m.%Y')
        d2 = datetime.strptime(list_date[i+1], '%d.%m.%Y')
        delta = d2 - d1
        for j in range(delta.days + 1):
            d = d1 + timedelta(j)
            if d.weekday() in weekday:
                list_days.append(d.strftime('%d.%m'+year))
                if d.weekday() in doubleday:
                    list_days.append(d.strftime('%d.%m'+year))
            if d.weekday() in list_all_days:
                all_days += 1
            dict_days[d.weekday()] += 1

def checkValidDatesArr(d1, d2):
    d1 = datetime.strptime(d1, '%d.%m.%Y')
    d2 = datetime.strptime(d2, '%d.%m.%Y')
    delta = d2 - d1
    if delta.days <=0:
        return False
    else:
        return True

def readWeekDays(err = False):
    global week_days, double_days
    week_days = []
    double_days = []
    frame_2.setStyleSheet('QWidget {}')
    pushButtonFill.setText('Заполнить')
    pushButtonFill.setStyleSheet('QWidget {font-size: 18pt;}')
    labelProgress.setText('Заполнено 0 из '+str(count_hours))
    progressBar.setValue(0)
    for i in range(6):
        if globals()['checkBoxWeek'+str(i)].isChecked():
            week_days.append(i)
        if globals()['spinBoxWeek'+str(i)].value() == 2:
            double_days.append(i)
    if not err:
        makeListOfDays(week_days, True, double_days)
    label_selected_days.setText('Выбрано: '+str(len(list_days)))
    plainTextEdit.setPlainText('\n'.join(list_days))
    if file_name:
        if count_hours == len(list_days) and count_hours != 0:
            label_selected_days.setStyleSheet('QWidget {color: rgb(28, 153, 0);%s}'%WIN_css)
        else:
            label_selected_days.setStyleSheet('QWidget {color: rgb(255, 0, 0);%s}'%WIN_css)

class ReadHoursThread(QtCore.QThread):
    labelHoursChangeSignal = QtCore.pyqtSignal(str)
    def  __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
    def run(self):
        global count_hours, one, column_with_days, list_one
        table = getTable()
        count_hours = 0
        if table:
            one = 2
            find = False
            rows = table.rows
            for i in range(1, 5):
                row = rows[i].cells
                if find:
                    break
                for j in range(one, len(row)):
                    if row[j].text.strip() == '1':
                        one = j
                        column_with_days = one + 1
                        find = True
                        break
            list_one = []
            columns = table.columns[one].cells
            for cell in columns:
                list_one.append(0)
                if cell.text.strip() == '1':
                    list_one[-1] = 1
                    count_hours += 1
                    self.labelHoursChangeSignal.emit('Найдено часов: '+str(count_hours))
                    sleep(0.001)
            if 1 not in list_one:
                one = 0
                column_with_days = 1
                list_one = []
                columns = table.columns[one].cells
                for i in range(len(columns)):
                    list_one.append(0)
                    if columns[i].text.strip() != '':
                        if columns[i].text.strip()[0].isdigit():
                            if columns[i].text.strip() != table.columns[one+1].cells[i].text.strip():
                                list_one[-1] = 1
                                count_hours += 1
                                self.labelHoursChangeSignal.emit('Найдено часов: '+str(count_hours))
                                sleep(0.001)
            self.labelHoursChangeSignal.emit('Найдено часов: '+str(count_hours))
        else:
            self.labelHoursChangeSignal.emit('Таблица не найдена')

def on_finished_read_hours():
    pushButtonFill.setEnabled(True)
    pushButton.setEnabled(True)
    comboBoxTables.setEnabled(True)
    comboBoxColumns.setEnabled(True)
    labelProgress.setText('Заполнено 0 из '+str(count_hours))
    progressBar.setValue(0)
    if count_hours % 34 == 0 and count_hours != 0:
        labelHours.setStyleSheet('QWidget {color: rgb(28, 153, 0);%s}'%WIN_css)
    else:
        labelHours.setStyleSheet('QWidget {color: rgb(255, 0, 0);%s}'%WIN_css)
        labelHours.setText('Найдено часов: '+str(count_hours)+'\nПроверьте столбец "Количество часов"')
    if count_hours == len(list_days) and count_hours != 0:
        label_selected_days.setStyleSheet('QWidget {color: rgb(28, 153, 0);%s}'%WIN_css)
    else:
        label_selected_days.setStyleSheet('QWidget {color: rgb(255, 0, 0);%s}'%WIN_css)
    if table_number != -1:
        win.comboBoxColumns.addItems(getColumnsNames())
    win.comboBoxColumns.setCurrentIndex(column_with_days)
    if win.comboBoxTables.count() == 0:
        win.comboBoxTables.addItems(getListOfTables())

def on_started_read_hours():
    pushButtonFill.setEnabled(False)
    pushButton.setEnabled(False)
    comboBoxTables.setEnabled(False)
    comboBoxColumns.setEnabled(False)
    labelHours.setStyleSheet('QWidget {%s}'%WIN_css)
    label_selected_days.setStyleSheet('QWidget {%s}'%WIN_css)
    labelHours.setText('Найдено часов: 0')
    win.comboBoxColumns.clear()
    win.resize(100, 100)

def on_label_hours_change(s):
    labelHours.setText(s)

read_hours_thread = ReadHoursThread()
read_hours_thread.finished.connect(on_finished_read_hours)
read_hours_thread.started.connect(on_started_read_hours)
read_hours_thread.labelHoursChangeSignal.connect(on_label_hours_change, QtCore.Qt.QueuedConnection)

def openFiles():
    global file_name, table_number
    frame_3.setStyleSheet('QWidget {}')
    pushButtonFill.setText('Заполнить')
    pushButtonFill.setStyleSheet('QWidget {font-size: 18pt;}')
    label_files.setText('')
    progressBar.setValue(0)
    labelProgress.setText('')
    labelHours.setText('')
    label_selected_days.setStyleSheet('QWidget {%s}'%WIN_css)
    file_name = QtWidgets.QFileDialog.getOpenFileName(win, 'Выберите файл', '', '(*.docx)')[0]
    if file_name:
        label_files.setText(os.path.basename(file_name))
        table_number = -1
        win.comboBoxTables.clear()
        read_hours_thread.start()

class FillTableThread(QtCore.QThread):
    labelProgressChangeSignal = QtCore.pyqtSignal(str)
    progressBarChangeSignal = QtCore.pyqtSignal(int)
    def  __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
    def run(self):
        doc = docx.Document(file_name)
        tables = doc.tables
        table = tables[table_number]
        if table:
            filled = 0
            col = table.columns[column_with_days].cells
            for i in range(len(col)):
                if list_one[i] == 1:
                    col[i].text = list_days[filled]
                    filled += 1
                    self.labelProgressChangeSignal.emit('Заполнено '+str(filled)+' из '+str(count_hours))
                    self.progressBarChangeSignal.emit(filled*100//count_hours)
                    sleep(0.001)

            doc.save(file_name)

def on_label_progress_change(s):
    labelProgress.setText(s)

def on_rogress_bar_change(n):
    progressBar.setValue(n)

def on_finished_fill_table():
    for i in range(4):
        globals()['frame_'+str(i)].setEnabled(True)
    pushButtonFill.setEnabled(True)
    pushButtonFill.setText('Готово')
    pushButtonFill.setStyleSheet('QWidget {font-size: 18pt; color: rgb(28, 153, 0);}')

def on_started_fill_table():
    for i in range(4):
        globals()['frame_'+str(i)].setEnabled(False)
    pushButtonFill.setEnabled(False)

fill_table_thread = FillTableThread()
fill_table_thread.finished.connect(on_finished_fill_table)
fill_table_thread.started.connect(on_started_fill_table)
fill_table_thread.labelProgressChangeSignal.connect(on_label_progress_change, QtCore.Qt.QueuedConnection)
fill_table_thread.progressBarChangeSignal.connect(on_rogress_bar_change, QtCore.Qt.QueuedConnection)


def fill():
    global list_days
    pushButtonFill.setText('Заполнить')
    if len(list_days) == 0 or not file_name or abs(len(list_days)-count_hours) > 1:
        if len(list_days) == 0 or abs(len(list_days)-count_hours) > 1:
            frame_2.setStyleSheet('QWidget {background-color: rgb(255, 190, 191);}')
        if not file_name:
            frame_3.setStyleSheet('QWidget {background-color: rgb(255, 190, 191);}')
        return
    if count_hours == 0:
        return
    if len(list_days) < count_hours:
        list_days.append('')
    fill_table_thread.start()

def getTable():
    if table_number == -1:
        getListOfTables()
    if table_number == -1:
        return
    doc = docx.Document(file_name)
    tables = doc.tables
    return(tables[table_number])

def getListOfTables():
    global table_number
    list_tables = []
    doc = docx.Document(file_name)
    tables = doc.tables
    for i in range(len(tables)):
        if len(tables[i].rows) > 34:
            if table_number == -1:
                table_number = i
                return
            list_tables.append('Таблица '+str(i+1))
    return(list_tables)

def getColumnsNames():
    list_of_columns = []
    doc = docx.Document(file_name)
    tables = doc.tables
    line = 0
    rows = [tables[table_number].rows[0].cells, tables[table_number].rows[1].cells]
    if rows[0][0].text == rows[1][0].text:
        line = 1
    i = 0
    for cell in rows[line]:
        s = cell.text.replace('\n', ' ')
        if line == 1:
            s2 = rows[0][i].text.replace('\n', ' ')
            if s != s2:
                s = s2 + ' ' + s
        i += 1
        list_of_columns.append(s)
    return(list_of_columns)

pushButtonFill.setFocus()
fixFontIfWIN()
loadDateFromFile()
loadDateToApp()
makeListOfDays()
loadWeekDays()
connectSignals()

win.show()
sys.exit(app.exec())
