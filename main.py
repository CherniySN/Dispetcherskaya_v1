from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QGridLayout, QWidget, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import QDate
import pickle
from PyQt5.QtCore import QSize, Qt
import pyodbc

Form, Window = uic.loadUiType("MaineWindow.ui")
app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

form.comboBox_14.addItems(["Белый", "Черный", "Синий", "Красный", "Желтый", "Зеленый", "Бежевый", "Оранжевый"])
form.comboBox_15.addItems(["Хорошее", "Удовлетворительное", "Плохое"])
form.comboBox_16.addItems(["1", "2", "3", "4", "5"])
form.comboBox.addItems(["Назначена", "В работе", "Выполнена"])
form.comboBox_10.addItems(["Назначена", "В работе", "Выполнена"])
# Звполняем табльцу на вкладки Магазины спец.техники
form.tableWidget_5.setColumnCount(6)  # Устанавливаем 6 колонк
form.tableWidget_5.setRowCount(100)  # и 1000 строку в таблице
form.tableWidget_5.setHorizontalHeaderLabels(
    ["Модель", "Год", "Цена", "Количество",
     "Название диллерского центра", "Телефон диллерского центра"])  # Устанавливаем заголовки таблицы
# Устанавливаем выравнивание на заголовки
form.tableWidget_5.horizontalHeaderItem(0).setTextAlignment(Qt.AlignLeft)
form.tableWidget_5.horizontalHeaderItem(1).setTextAlignment(Qt.AlignLeft)
form.tableWidget_5.horizontalHeaderItem(2).setTextAlignment(Qt.AlignLeft)
form.tableWidget_5.horizontalHeaderItem(3).setTextAlignment(Qt.AlignLeft)
form.tableWidget_5.horizontalHeaderItem(4).setTextAlignment(Qt.AlignLeft)
form.tableWidget_5.horizontalHeaderItem(5).setTextAlignment(Qt.AlignLeft)
# делаем ресайз колонок по содержимому
form.tableWidget_5.resizeColumnsToContents()

# Звполняем табльцу на вкладки Магазины спец.техники
form.tableWidget_6.setColumnCount(8)  # Устанавливаем 6 колонк
form.tableWidget_6.setRowCount(100)  # и 1000 строку в таблице
form.tableWidget_6.setHorizontalHeaderLabels(
    ["Фамилия клиента", "Имя клиента", "Причина обращения", "Статус",
     "Наименование работ", "Цена работ", "Позывной сотрудника",
     "Гос.номер спец.средства"])  # Устанавливаем заголовки таблицы
# Устанавливаем выравнивание на заголовки
form.tableWidget_6.horizontalHeaderItem(0).setTextAlignment(Qt.AlignLeft)
form.tableWidget_6.horizontalHeaderItem(1).setTextAlignment(Qt.AlignLeft)
form.tableWidget_6.horizontalHeaderItem(2).setTextAlignment(Qt.AlignLeft)
form.tableWidget_6.horizontalHeaderItem(3).setTextAlignment(Qt.AlignLeft)
form.tableWidget_6.horizontalHeaderItem(4).setTextAlignment(Qt.AlignLeft)
form.tableWidget_6.horizontalHeaderItem(5).setTextAlignment(Qt.AlignLeft)
form.tableWidget_6.horizontalHeaderItem(6).setTextAlignment(Qt.AlignLeft)
form.tableWidget_6.horizontalHeaderItem(7).setTextAlignment(Qt.AlignLeft)
# делаем ресайз колонок по содержимому
form.tableWidget_6.resizeColumnsToContents()
# Звполняем табльцу на вкладки Автомобили
form.tableWidget_2.setColumnCount(5)  # Устанавливаем 5 колонк
form.tableWidget_2.setRowCount(100)  # и 1000 строку в таблице
form.tableWidget_2.setHorizontalHeaderLabels(
    ["Модель", "Цвет", "Гос.номер", "Дата выпуска", "Техническое состояние"])  # Устанавливаем заголовки таблицы
# Устанавливаем выравнивание на заголовки
form.tableWidget_2.horizontalHeaderItem(0).setTextAlignment(Qt.AlignLeft)
form.tableWidget_2.horizontalHeaderItem(1).setTextAlignment(Qt.AlignLeft)
form.tableWidget_2.horizontalHeaderItem(2).setTextAlignment(Qt.AlignLeft)
form.tableWidget_2.horizontalHeaderItem(3).setTextAlignment(Qt.AlignLeft)
form.tableWidget_2.horizontalHeaderItem(4).setTextAlignment(Qt.AlignLeft)
# делаем ресайз колонок по содержимому
form.tableWidget_2.resizeColumnsToContents()

# Звполняем табльцу на вкладки Автомобили, магазин
form.tableWidget_3.setColumnCount(5)  # Устанавливаем 6 колонк
form.tableWidget_3.setRowCount(100)  # и 1000 строку в таблице
form.tableWidget_3.setHorizontalHeaderLabels(
    ["Модель", "Год выпуска", "Мощьность двигателя", "Цена",
     "Количество шт.на складе"])  # Устанавливаем заголовки таблицы
# Устанавливаем выравнивание на заголовки
form.tableWidget_3.horizontalHeaderItem(0).setTextAlignment(Qt.AlignLeft)
form.tableWidget_3.horizontalHeaderItem(1).setTextAlignment(Qt.AlignLeft)
form.tableWidget_3.horizontalHeaderItem(2).setTextAlignment(Qt.AlignLeft)
form.tableWidget_3.horizontalHeaderItem(3).setTextAlignment(Qt.AlignLeft)
form.tableWidget_3.horizontalHeaderItem(4).setTextAlignment(Qt.AlignLeft)
# делаем ресайз колонок по содержимому
form.tableWidget_3.resizeColumnsToContents()

# Звполняем табльцу на вкладки Автомобили, магазин
form.tableWidget_4.setColumnCount(6)  # Устанавливаем 6 колонк
form.tableWidget_4.setRowCount(100)  # и 1000 строку в таблице
form.tableWidget_4.setHorizontalHeaderLabels(
    ["Фимилия", "Имя", "Отчество", "Дата рождения", "Разряд", "Позывной"])  # Устанавливаем заголовки таблицы
# Устанавливаем выравнивание на заголовки
form.tableWidget_4.horizontalHeaderItem(0).setTextAlignment(Qt.AlignLeft)
form.tableWidget_4.horizontalHeaderItem(1).setTextAlignment(Qt.AlignLeft)
form.tableWidget_4.horizontalHeaderItem(2).setTextAlignment(Qt.AlignLeft)
form.tableWidget_4.horizontalHeaderItem(3).setTextAlignment(Qt.AlignLeft)
form.tableWidget_4.horizontalHeaderItem(4).setTextAlignment(Qt.AlignLeft)
form.tableWidget_4.horizontalHeaderItem(5).setTextAlignment(Qt.AlignLeft)
# делаем ресайз колонок по содержимому
form.tableWidget_4.resizeColumnsToContents()

# Звполняем табльцу на вкладки Адреса
form.tableWidget.setColumnCount(4)  # Устанавливаем 6 колонк
form.tableWidget.setRowCount(100)  # и 1000 строку в таблице
form.tableWidget.setHorizontalHeaderLabels(
    ["Город", "Улица", "Дом", "Квартира"])  # Устанавливаем заголовки таблицы
# Устанавливаем выравнивание на заголовки
form.tableWidget.horizontalHeaderItem(0).setTextAlignment(Qt.AlignLeft)
form.tableWidget.horizontalHeaderItem(1).setTextAlignment(Qt.AlignLeft)
form.tableWidget.horizontalHeaderItem(2).setTextAlignment(Qt.AlignLeft)
form.tableWidget.horizontalHeaderItem(3).setTextAlignment(Qt.AlignLeft)
# делаем ресайз колонок по содержимому
form.tableWidget.resizeColumnsToContents()
# ---------------------------- Заполняем табличку -------------------------------------------------------
connection_to_db = pyodbc.connect(
    r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
cursor = connection_to_db.cursor()
cursor.execute('select City, Street, House, Flat from Adress;')
i = 0
while 1:
    row = cursor.fetchone()
    if not row:
        break

    form.tableWidget.setItem(i, 0, QTableWidgetItem(row.City))
    form.tableWidget.setItem(i, 1, QTableWidgetItem(row.Street))
    form.tableWidget.setItem(i, 2, QTableWidgetItem(str(row.House)))
    form.tableWidget.setItem(i, 3, QTableWidgetItem(str(row.Flat)))
    i = i + 1

# делаем ресайз колонок по содержимому
form.tableWidget.resizeColumnsToContents()
connection_to_db.close()
# -------------------------------------------------------------------------------------------------------
# ---------------------------- Заполняем выпадающий список -------------------------------------------------------
connection_to_db = pyodbc.connect(
    r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
cursor = connection_to_db.cursor()
cursor.execute('select Car.Government_number from Car;')
while True:
    row3 = cursor.fetchone()
    if not row3:
        break
    form.comboBox_11.addItems([row3.Government_number])  # Добавление в выпод. список
connection_to_db.close()
# -------------------------------------------------------------------------------------------------------
# ---------------------------- Заполняем выпадающий список -------------------------------------------------------
connection_to_db = pyodbc.connect(
    r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
cursor = connection_to_db.cursor()
cursor.execute('Select Telephone from Client;')
while True:
    row3 = cursor.fetchone()
    if not row3:
        break
    form.comboBox_8.addItems([str(row3.Telephone)])  # Добавление в выпод. список
connection_to_db.close()
# -------------------------------------------------------------------------------------------------------
# ---------------------------- Заполняем выпадающий список покупки автомобилей----------------------------
connection_to_db = pyodbc.connect(
    r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
cursor = connection_to_db.cursor()
cursor.execute("SELECT * FROM OPENQUERY(POSTGRESQL, 'select idauto from Auto;');")
while True:
    row3 = cursor.fetchone()
    if not row3:
        break
    form.comboBox_13.addItems([str(row3.idauto)])  # Добавление в выпод. список
connection_to_db.close()
# -------------------------------------------------------------------------------------------------------
# ---------------------------- Заполняем выпадающий список Работы -------------------------------------------------------
connection_to_db = pyodbc.connect(
    r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
cursor = connection_to_db.cursor()
cursor.execute('SELECT Name_of_works FROM TypeOfWork;')
while True:
    row3 = cursor.fetchone()
    if not row3:
        break
    form.comboBox_9.addItems([row3.Name_of_works])  # Добавление в выпод. список
    form.comboBox_2.addItems([row3.Name_of_works])
connection_to_db.close()
# -------------------------------------------------------------------------------------------------------
# ---------------------------- Заполняем выпадающий список Сотрудники--------------------------------------
connection_to_db = pyodbc.connect(
    r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
cursor = connection_to_db.cursor()
cursor.execute('SELECT Employee_code FROM Employee;')
while True:
    row3 = cursor.fetchone()
    if not row3:
        break
    form.comboBox_4.addItems([str(row3.Employee_code)])  # Добавление в выпод. список
    form.comboBox_7.addItems([str(row3.Employee_code)])
connection_to_db.close()
# -----------------------------------------------------------------------------------------------------------
# ---------------------------- Заполняем выпадающий список -------------------------------------------------------
connection_to_db = pyodbc.connect(
    r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
cursor = connection_to_db.cursor()
cursor.execute('select Address_code from Adress;')
while True:
    row3 = cursor.fetchone()
    if not row3:
        break
    form.comboBox_12.addItems([str(row3.Address_code)])  # Добавление в выпод. список
    form.comboBox_6.addItems([str(row3.Address_code)])  # Добавление в выпод. список
    form.comboBox_5.addItems([str(row3.Address_code)])  # Добавление в выпод. список
    form.comboBox_3.addItems([str(row3.Address_code)])  # Добавление в выпод. список

connection_to_db.close()


def on_click_reload_adres():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        cursor = connection_to_db.cursor()
        cursor.execute('select City, Street, House, Flat from Adress;')
        i = 0
        while 1:
            row = cursor.fetchone()
            if not row:
                break

            form.tableWidget.setItem(i, 0, QTableWidgetItem(row.City))
            form.tableWidget.setItem(i, 1, QTableWidgetItem(row.Street))
            form.tableWidget.setItem(i, 2, QTableWidgetItem(str(row.House)))
            form.tableWidget.setItem(i, 3, QTableWidgetItem(str(row.Flat)))
            i = i + 1

        # делаем ресайз колонок по содержимому
        form.tableWidget.resizeColumnsToContents()
    except:
        print("Что-то пошло не так при подключении к MS SQL SERVER")
    connection_to_db.close()  # закрываем связь с БД


# ---Ищем адрес в табличке
def on_click_finde_adres():
    try:
        finde = form.plainTextEdit.toPlainText()
        items = form.tableWidget.findItems('%s' % finde, Qt.MatchFixedString)
        if items:
            form.tableWidget.setCurrentItem(items[1])
            form.label_2.setText("СТАТУС: Адрес найден")
        else:
            form.label_2.setText("СТАТУС: Адрес не найден")
    except:
        form.label_2.setText("СТАТУС: Что-то пошло не так попробуйте сново")


# ------------Заведение нового адреса
def on_click_new_adres():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        City = form.plainTextEdit_2.toPlainText()
        Street = form.plainTextEdit_3.toPlainText()
        House = form.plainTextEdit_4.toPlainText()
        Flat = form.plainTextEdit_5.toPlainText()
        Str_House = int(House)
        Str_Flat = int(Flat)
        cursor = connection_to_db.cursor()
        cursor.execute(
            "Insert INTO Adress([City],[Street],[House],[Flat]) VALUES (?,?,?,?)",
            City, Street, Str_House, Str_Flat)
        cursor.commit()
    except:
        print("Что-то пошло не так в БД при заведении студента")
    connection_to_db.close()


def on_click_store_of_stuff():
    try:
        connection_to_db = pyodbc.connect(  # соединяемся с БД
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        cursor = connection_to_db.cursor()
        cursor.execute("SELECT * FROM modelpricestore;")  # запрос с использованием представления
        i = 0
        while True:
            row1 = cursor.fetchone()  # получаем данные построчно из запроса
            if not row1:
                break
            form.tableWidget_5.setItem(i, 0, QTableWidgetItem(str(row1.model)))
            form.tableWidget_5.setItem(i, 1, QTableWidgetItem(str(row1.yearofauto)))
            form.tableWidget_5.setItem(i, 2, QTableWidgetItem(str(row1.price)))
            form.tableWidget_5.setItem(i, 3, QTableWidgetItem(str(row1.valueofauto)))
            form.tableWidget_5.setItem(i, 4, QTableWidgetItem(str(row1.namestore)))
            form.tableWidget_5.setItem(i, 5, QTableWidgetItem(str(row1.telstore)))
            i = i + 1
            # делаем ресайз колонок по содержимому
            form.tableWidget_5.resizeColumnsToContents()
    except:
        form.label_57.setText("СТАТУС:  Что-то пошло не так в БД")
        print("Что-то пошло не так при подключении к MS SQL SERVER")
    connection_to_db.close()


def on_click_reload_car():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        cursor = connection_to_db.cursor()
        cursor.execute(
            'SELECT Car.Car_model, Car.Color, Car.Government_number, Car.Year_of_release, Car.Technical_condition FROM Car;')
        i = 0
        while True:
            row = cursor.fetchone()
            if not row:
                break
            form.tableWidget_2.setItem(i, 0, QTableWidgetItem(row.Car_model))
            form.tableWidget_2.setItem(i, 1, QTableWidgetItem(row.Color))
            form.tableWidget_2.setItem(i, 2, QTableWidgetItem(row.Government_number))
            form.tableWidget_2.setItem(i, 3, QTableWidgetItem(str(row.Year_of_release)))
            form.tableWidget_2.setItem(i, 4, QTableWidgetItem(row.Technical_condition))
            i = i + 1

        # делаем ресайз колонок по содержимому
        form.tableWidget_2.resizeColumnsToContents()
    except:
        print("Что-то пошло не так при подключении к MS SQL SERVER")
    connection_to_db.close()  # закрываем связь с БД


# -----Зполнение таблицы автомобилей,магазин кнопка обновить
def on_click_reload_carstore():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        cursor = connection_to_db.cursor()
        cursor.execute('SELECT * FROM modelpricestoremodel;')
        i = 0
        while True:
            row = cursor.fetchone()
            if not row:
                break
            form.tableWidget_3.setItem(i, 0, QTableWidgetItem(row.model))
            form.tableWidget_3.setItem(i, 1, QTableWidgetItem(str(row.yearofauto)))
            form.tableWidget_3.setItem(i, 2, QTableWidgetItem(str(row.powerofauto)))
            form.tableWidget_3.setItem(i, 3, QTableWidgetItem(str(row.price)))
            form.tableWidget_3.setItem(i, 4, QTableWidgetItem(str(row.valueofauto)))
            i = i + 1

        # делаем ресайз колонок по содержимому
        form.tableWidget_3.resizeColumnsToContents()
    except:
        print("Что-то пошло не так при подключении к MS SQL SERVER")
    connection_to_db.close()  # закрываем связь с БД


# ---Ищем авто в табличке
def on_click_finde_car():
    try:
        finde = form.plainTextEdit_7.toPlainText()
        items = form.tableWidget_2.findItems('%s' % finde, Qt.MatchFixedString)
        if items:
            form.tableWidget_2.setCurrentItem(items[0])
            form.label_11.setText("СТАТУС: Спец.стредство найден")
        else:
            form.label_11.setText("СТАТУС: Спец.стредство не найден")
    except:
        form.label_11.setText("СТАТУС: Что-то пошло не так, попробуйте снова")


def on_click_buy_carstore():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        finde = form.comboBox_13.currentText()
        Intfinde = int(finde)
        cursor = connection_to_db.cursor()
        cursor.execute("SELECT * FROM OPENQUERY(POSTGRESQL, 'select model from Auto where idauto = %s;')" % Intfinde)
        row1 = cursor.fetchone()
        Car_model = row1[0]
        Government_number = form.plainTextEdit_8.toPlainText()
        Color = form.comboBox_14.currentText()
        Year_of_release = '2067-12-12'
        Technical_condition = form.comboBox_15.currentText()
        cursor = connection_to_db.cursor()
        cursor.execute(
            "INSERT into Car(Car_model,Government_number,Color,Year_of_release,Technical_condition) VALUES (?,?,?,?,?);",
            Car_model, Government_number, Color, Year_of_release, Technical_condition)
        cursor.commit()
    except:
        print("Ну е мое. Что-то пошло не так")
        connection_to_db.close()  # закрываем связь с БД


def on_click_reload_employe():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        cursor = connection_to_db.cursor()
        cursor.execute(
            'select [Employee_surname],[Employee_name],[Employee_patronymic],[Employee_date_of_birth],[Category],[Callsign] from Employee;')
        i = 0
        while True:
            row = cursor.fetchone()
            if not row:
                break
            form.tableWidget_4.setItem(i, 0, QTableWidgetItem(row.Employee_surname))
            form.tableWidget_4.setItem(i, 1, QTableWidgetItem(row.Employee_name))
            form.tableWidget_4.setItem(i, 2, QTableWidgetItem(row.Employee_patronymic))
            form.tableWidget_4.setItem(i, 3, QTableWidgetItem(str(row.Employee_date_of_birth)))
            form.tableWidget_4.setItem(i, 4, QTableWidgetItem(str(row.Category)))
            form.tableWidget_4.setItem(i, 5, QTableWidgetItem(row.Callsign))
            i = i + 1

        # делаем ресайз колонок по содержимому
        form.tableWidget_4.resizeColumnsToContents()

    except:
        form.label_20.setText("СТАТУС:  Что-то пошло не так в БД")
        print("Что-то пошло не так при подключении к MS SQL SERVER")
    connection_to_db.close()  # закрываем связь с БД


# ---Ищем авто в табличке
def on_click_finde_employe():
    try:
        finde = form.plainTextEdit_12.toPlainText()
        items = form.tableWidget_4.findItems('%s' % finde, Qt.MatchFixedString)
        if items:
            form.tableWidget_4.setCurrentItem(items[0])
            form.label_20.setText("СТАТУС: Сотрудник найден")
        else:
            form.label_20.setText("СТАТУС: Сотрудник не найден")
    except:
        form.label_20.setText("СТАТУС: Что-то пошло не так, попробуйте снова")


def on_click_save_newemploye():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        Employee_surname = form.plainTextEdit_15.toPlainText()
        Employee_name = form.plainTextEdit_11.toPlainText()
        Employee_patronymic = form.plainTextEdit_14.toPlainText()
        Employee_date_of_birth = form.dateEdit.dateTime().toString('yyyy-MM-dd')
        Category = form.comboBox_16.currentText()
        Int_Category = int(Category)
        Callsign = form.plainTextEdit_17.toPlainText()
        Vehicle_code = form.comboBox_11.currentText()
        cursor = connection_to_db.cursor()
        cursor.execute("SELECT [Vehicle_code] FROM [Car] WHERE [Government_number] = '%s';" % Vehicle_code)
        row = cursor.fetchone()
        Int_Vehicle_code = int(row[0])
        Address_code = form.comboBox_12.currentText()
        Int_Address_code = int(Address_code)
        cursor = connection_to_db.cursor()
        cursor.execute(
            "INSERT into Employee([Employee_surname],[Employee_name],[Employee_patronymic],[Employee_date_of_birth],[Category],[Callsign],[Vehicle_code],[Address_code]) VALUES (?, ?, ?, ?,?,?,?,?);",
            Employee_surname, Employee_name, Employee_patronymic, Employee_date_of_birth, Int_Category, Callsign,
            Int_Vehicle_code, Int_Address_code)
        cursor.commit()
    except:
        form.label_20.setText("СТАТУС:  Что-то пошло не так в БД")
        print("Что-то пошло не так при подключении к MS SQL SERVER")
    connection_to_db.close()  # закрываем связь с БД


def on_click_reload_request():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        cursor = connection_to_db.cursor()
        cursor.execute(
            'SELECT DISTINCT Client.Client_surname, Client.Client_name,Request.Initial_value,Request.Application_status, TypeOfWork.Name_of_works, TypeOfWork.Price, Employee.Callsign, Car.Government_number FROM Client, TypeOfWork, Request, Employee,Car where Client.Client_code = Request.Client_code and Request.Work_code = TypeOfWork.Work_code and Request.Employee_code = Employee.Employee_code and Employee.Vehicle_code = Car.Vehicle_code;')
        i = 0
        while True:
            row = cursor.fetchone()
            if not row:
                break
            form.tableWidget_6.setItem(i, 0, QTableWidgetItem(row.Client_surname))
            form.tableWidget_6.setItem(i, 1, QTableWidgetItem(row.Client_name))
            form.tableWidget_6.setItem(i, 2, QTableWidgetItem(row.Initial_value))
            form.tableWidget_6.setItem(i, 3, QTableWidgetItem(row.Application_status))
            form.tableWidget_6.setItem(i, 4, QTableWidgetItem(row.Name_of_works))
            form.tableWidget_6.setItem(i, 5, QTableWidgetItem(str(row.Price)))
            form.tableWidget_6.setItem(i, 6, QTableWidgetItem(row.Callsign))
            form.tableWidget_6.setItem(i, 7, QTableWidgetItem(row.Government_number))
            i = i + 1

        # делаем ресайз колонок по содержимому
        form.tableWidget_6.resizeColumnsToContents()
    except:
        print("Что-то пошло не так при подключении к MS SQL SERVER")
    connection_to_db.close()  # закрываем связь с БД


def on_click_new_reqest_client():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        Initial_value = form.plainTextEdit_40.toPlainText()
        End_value = form.plainTextEdit_34.toPlainText()
        The_date_of_the_beginning = '2051-07-10'
        Start_time = '11:00:00'
        Expiration_date = '2051-07-10'
        End_time = '11:00:00'
        Application_status = form.comboBox_10.currentText()
        finde_id_worke = form.comboBox_9.currentText()
        cursor = connection_to_db.cursor()
        cursor.execute(
            "SELECT TypeOfWork.Work_code FROM TypeOfWork WHERE TypeOfWork.Name_of_works = '%s';" % finde_id_worke)
        row1 = cursor.fetchone()
        ID_Work_code = int(row1[0])
        Address_code = form.comboBox_6.currentText()
        Int_Address_code = int(Address_code)
        Employee_code = form.comboBox_7.currentText()
        Int_Employee_code = int(Employee_code)
        finde_Client_code = form.comboBox_8.currentText()
        Int_finde_Client_code = int(finde_Client_code)
        cursor = connection_to_db.cursor()
        cursor.execute("SELECT Client.Client_code FROM Client WHERE Client.Telephone =  %s;" % Int_finde_Client_code)
        row1 = cursor.fetchone()
        ID_Client = int(row1[0])
        cursor = connection_to_db.cursor()
        cursor.execute(
            "INSERT into Request(Initial_value,End_value,The_date_of_the_beginning,Start_time,Expiration_date,End_time,Application_status,Work_code,Address_code,Employee_code,Client_code) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            Initial_value, End_value, The_date_of_the_beginning, Start_time, Expiration_date, End_time,
            Application_status, ID_Work_code, Int_Address_code, Int_Employee_code, ID_Client)
        cursor.commit()  # Обратите внимание на вызовы cnxn.commit(). Вы должны позвонить commit, иначе ваши изменения будут потеряны! Когда соединение будет закрыто,все ожидающие изменения будут откатаны.

    except:
        print("Ну е мое.")
        connection_to_db.close()  # закрываем связь с БД


def on_click_reqest_client():
    try:
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        Client_surname = form.plainTextEdit_21.toPlainText()
        Client_name = form.plainTextEdit_24.toPlainText()
        Client_patronymic = form.plainTextEdit_25.toPlainText()
        Date_of_Birth = form.dateEdit_2.dateTime().toString('yyyy-MM-dd')
        Telephone = form.plainTextEdit_22.toPlainText()
        Int_Telephone = int(Telephone)
        Address_code = form.comboBox_5.currentText()
        Int_Address_code = int(Address_code)
        cursor = connection_to_db.cursor()
        cursor.execute(
            "INSERT into Client(	Client_surname,Client_name,Client_patronymic,Date_of_Birth,Telephone,Address_code) VALUES (?, ?, ?, ?, ?, ?)",
            Client_surname, Client_name, Client_patronymic, Date_of_Birth, Int_Telephone, Int_Address_code)
        cursor.commit()
        connection_to_db = pyodbc.connect(
            r'Driver={SQL Server};Server=DESKTOP-EPARK0G\SQLEXPRESS;Database=DispatchDepartment;Trusted_Connection=yes;')
        Initial_value = form.plainTextEdit_33.toPlainText()
        End_value = form.plainTextEdit_27.toPlainText()
        The_date_of_the_beginning = '2051-07-10'
        Start_time = '11:00:00'
        Expiration_date = '2051-07-10'
        End_time = '11:00:00'
        Application_status = form.comboBox.currentText()
        finde_id_worke = form.comboBox_2.currentText()
        cursor = connection_to_db.cursor()
        cursor.execute(
            "SELECT TypeOfWork.Work_code FROM TypeOfWork WHERE TypeOfWork.Name_of_works = '%s';" % finde_id_worke)
        row1 = cursor.fetchone()
        ID_Work_code = int(row1[0])
        Address_code = form.comboBox_3.currentText()
        Int_Address_code = int(Address_code)
        Employee_code = form.comboBox_4.currentText()
        Int_Employee_code = int(Employee_code)
        cursor = connection_to_db.cursor()
        cursor.execute("SELECT Client.Client_code FROM Client WHERE Client.Telephone =  %s;" % Int_Telephone)
        row1 = cursor.fetchone()
        ID_Client = int(row1[0])
        cursor = connection_to_db.cursor()
        cursor.execute(
            "INSERT into Request(Initial_value,End_value,The_date_of_the_beginning,Start_time,Expiration_date,End_time,Application_status,Work_code,Address_code,Employee_code,Client_code) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            Initial_value, End_value, The_date_of_the_beginning, Start_time, Expiration_date, End_time,
            Application_status, ID_Work_code, Int_Address_code, Int_Employee_code, ID_Client)
        cursor.commit()  # Обратите внимание на вызовы cnxn.commit(). Вы должны позвонить commit, иначе ваши изменения будут потеряны! Когда соединение будет закрыто,все ожидающие изменения будут откатаны.

    except:
        print("Ну е мое. Что-то пошло не так")
        connection_to_db.close()  # закрываем связь с БД


form.pushButton_11.clicked.connect(on_click_store_of_stuff)
form.pushButton_2.clicked.connect(on_click_reload_adres)
form.pushButton.clicked.connect(on_click_finde_adres)
form.pushButton_3.clicked.connect(on_click_new_adres)
form.pushButton_5.clicked.connect(on_click_finde_car)
form.pushButton_4.clicked.connect(on_click_reload_car)
form.pushButton_7.clicked.connect(on_click_reload_carstore)
form.pushButton_6.clicked.connect(on_click_buy_carstore)
form.pushButton_8.clicked.connect(on_click_reload_employe)
form.pushButton_9.clicked.connect(on_click_finde_employe)
form.pushButton_10.clicked.connect(on_click_save_newemploye)
form.pushButton_14.clicked.connect(on_click_reload_request)
form.pushButton_13.clicked.connect(on_click_new_reqest_client)
form.pushButton_12.clicked.connect(on_click_reqest_client)

app.exec_()
