# @CODE_CO :) | TELEGRAM: @CODE_CO | GITHUB: CODE-CO |----------------------------------------------
#---------------------------------------------------------------------------------------------------
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem
import xlsxwriter, requests, sqlite3, vobject, time, sys, re
from PyQt5 import QtWidgets, QtCore
from bs4 import BeautifulSoup
from PyQt5.uic import loadUi
#---------------------------------------------------------------------------------------------------
CONNECT_SQLITE3 = sqlite3.connect('CONTACTS.DB')
CURSOR = CONNECT_SQLITE3.cursor()

for TABLE in ["BOTH", "MEN", "WOMEN"]: CURSOR.execute('''CREATE TABLE IF NOT EXISTS "{}" (id INTEGER PRIMARY KEY,name TEXT,phone TEXT,UNIQUE(name, phone))'''.format(TABLE))

DATABASE_START = [CURSOR.execute("SELECT NAME, PHONE FROM {}".format(TABLE)).fetchall() for TABLE in ["BOTH", "MEN", "WOMEN"]]

DATABASE = list()
RE_NAME = r'<strong class=\"text-lg font-bold text-gray-800\">\s*(.*?)\s*</strong>'
RE_NUMBER = r'(?<=تلفن همراه:</span>\n<span class=\"text-gray-700\">)[۰-۹]+'
MAIN_STYLE_SHEET = """QWidget#centralwidget{background-color: rgb(245, 221, 123)}"""
DARK_MODE_STYLE_SHEET = """QWidget#centralwidget{background-color: qlineargradient(spread:pad, x1:0.168421, y1:0.164773, x2:1, y2:1, stop:0 rgba(70, 170, 204, 255), stop:1 rgba(255, 255, 255, 255))}"""

GET = requests.Session()
#---------------------------------------------------------------------------------------------------
class Main(QMainWindow):
    def __init__(self):
        super(Main, self).__init__()
        loadUi("WINDOW1.UI", self)
        self.BUTTON_GET_DATA.clicked.connect(self.GET_DATA)
        self.BUTTON_VCF.clicked.connect(self.VCF)
        self.BUTTON_XLSX.clicked.connect(self.XLSX)
        self.B_DARK_MODE.clicked.connect(self.TOGGLE_DARK_MODE)
        self.BUTTON_WINDOW_2.clicked.connect(self.CHANGE_WINDOW)
        self.B_CLOSE.clicked.connect(self.CLOSE_APPLICATION)
        RUN_DATABASE(self)

    def CLOSE_APPLICATION(self):
        QApplication.quit()

    def XLSX(self):
        XLSX(self.NAMES, self.NUMBERS, self.RADIO_BUTTON)

    def VCF(self):
        VCF(self.NAMES, self.NUMBERS, self.RADIO_BUTTON)

    def CHANGE_WINDOW(self):
        WIDGET.setCurrentIndex(WIDGET.currentIndex()+1)

    def TOGGLE_DARK_MODE(self):
        DARK_MODE_FUNCTION(self)

    def GET_DATA(self):
        self.NAMES, self.NUMBERS = list(), list()
        self.SHOW.setText("0")
        self.TEXT_NAMES.setText("")
        self.TEXT_NUMBERS.setText("")
        self.TEXT_NAMES.setAlignment(QtCore.Qt.AlignCenter)
        self.TEXT_NUMBERS.setAlignment(QtCore.Qt.AlignCenter)
        CHECK_INTERNET_CONNECTION(self)
        WITCH_RADIO_BUTTON(self)
        self.BUTTON_GET_DATA.setEnabled(False)
        self.BUTTON_GET_DATA.setStyleSheet("""border-radius: 20px; color:rgb(177, 177, 177); font: 75 10pt "MS Shell Dlg 2"; background-color: rgb(85, 85, 85);""")
        PAGE = 0
        while True:
            PAGE += 1
            ELEMENTS = CONNECT(self.URL + str(PAGE))
            if NAMES_FOUND := [NAME for NAME in re.findall(RE_NAME, str(ELEMENTS)) if NAME != "علی اکبر پرهیزکار"]:
                NUMBERS_FOUND = re.findall(RE_NUMBER, str(ELEMENTS))
                self.NAMES.extend(NAMES_FOUND)
                self.NUMBERS.extend(NUMBERS_FOUND)
                self.TEXT_NAMES.append("\n".join(NAMES_FOUND))
                self.TEXT_NUMBERS.append("\n".join(NUMBERS_FOUND))
                self.SHOW.setText(str(len(self.NAMES)))
                APP.processEvents()
            else:
                break

        for ID, (NAME, NUMBER) in enumerate(zip(self.NAMES, self.NUMBERS), {"BOTH": len(DATABASE[0]), "MEN": len(DATABASE[1]), "WOMEN": len(DATABASE[2])}[self.RADIO_BUTTON]):
            ADD_CONTACTS(self, str(ID), NAME, NUMBER)

        RUN_DATABASE(self)
        self.BUTTON_GET_DATA.setEnabled(True)
        self.BUTTON_GET_DATA.setStyleSheet("""border-radius: 20px; color:rgb(255, 255, 255); font: 75 10pt "MS Shell Dlg 2"; background-color: rgb(85, 85, 85);""")
#---------------------------------------------------------------------------------------------------
class WINDOW_2(Main):
    def __init__(self):
        super(WINDOW_2, self).__init__()
        loadUi("WINDOW2.UI", self)
        self.tableWidget.verticalHeader().setDefaultAlignment(QtCore.Qt.AlignVCenter)
        self.B_LOAD.clicked.connect(self.LOAD_DATABASE)
        self.B_BACK.clicked.connect(self.CHANGE_WINDOW)
        self.B_DATABASE_AND_NEW.clicked.connect(self.CREATE_XLSX_VCF_FROM_BOTH)
        self.B_ONLY_DATABASE.clicked.connect(self.CREATE_XLSX_VCF_FROM_DATABASE)
        self.B_CLOSE.clicked.connect(self.CLOSE_APPLICATION)
        MAIN_WINDOW.B_DARK_MODE.clicked.connect(self.DARK_MODE_WINDOW_2)
        MAIN_WINDOW.BUTTON_WINDOW_2.clicked.connect(self.REFRESH_DATABASE_TAB)
        RUN_DATABASE(self)

    def CLOSE_APPLICATION(self):
        QApplication.quit()

    def DARK_MODE_WINDOW_2(self):
        DARK_MODE_FUNCTION(self)

    def CHANGE_WINDOW(self):
        WIDGET.setCurrentIndex(WIDGET.currentIndex()-1)

    def REFRESH_DATABASE_TAB(self):
        RUN_DATABASE(self)

    def CREATE_XLSX_VCF_FROM_BOTH(self):
        WITCH_RADIO_BUTTON(self)
        DATA_KEY = {"BOTH": DATABASE[0], "MEN": DATABASE[1], "WOMEN": DATABASE[2]}[self.RADIO_BUTTON]
        self.NAMES = [NAME[0] for NAME in DATA_KEY]
        self.NUMBERS = [NUMBER[1] for NUMBER in DATA_KEY]
        XLSX(self.NAMES, self.NUMBERS, self.RADIO_BUTTON)
        VCF(self.NAMES, self.NUMBERS, self.RADIO_BUTTON)

    def CREATE_XLSX_VCF_FROM_DATABASE(self):
        WITCH_RADIO_BUTTON(self)
        DATA_KEY = {"BOTH": DATABASE_START[0], "MEN": DATABASE_START[1], "WOMEN": DATABASE_START[2]}[self.RADIO_BUTTON]
        self.NAMES = [NAME[0] for NAME in DATA_KEY]
        self.NUMBERS = [NUMBER[1] for NUMBER in DATA_KEY]
        XLSX(self.NAMES, self.NUMBERS, self.RADIO_BUTTON)
        VCF(self.NAMES, self.NUMBERS, self.RADIO_BUTTON)

    def LOAD_DATABASE(self):
        self.tableWidget.clearContents()
        self.tableWidget.setRowCount(0)
        WITCH_RADIO_BUTTON(self)
        DATA_KEY = {"BOTH": DATABASE[0], "MEN": DATABASE[1], "WOMEN": DATABASE[2]}[self.RADIO_BUTTON]
        self.NAMES = [NAME[0] for NAME in DATA_KEY]
        self.NUMBERS = [NUMBER[1] for NUMBER in DATA_KEY]
        for NAME, NUMBER in zip(self.NAMES, self.NUMBERS):
            self.ADD_TO_TABLE(NAME, NUMBER)

    def ADD_TO_TABLE(self, NAME, NUMBER):
        ROW_POSITION = self.tableWidget.rowCount()
        self.tableWidget.insertRow(ROW_POSITION)

        NAME_ITEM = QTableWidgetItem(NAME)
        NAME_ITEM.setTextAlignment(QtCore.Qt.AlignCenter)
        self.tableWidget.setItem(ROW_POSITION, 0, NAME_ITEM)

        NUMBER_ITEM = QTableWidgetItem(NUMBER)
        NUMBER_ITEM.setTextAlignment(QtCore.Qt.AlignCenter)
        self.tableWidget.setItem(ROW_POSITION, 1, NUMBER_ITEM)
#---------------------------------------------------------------------------------------------------
def CONNECT(PAGE):
        return BeautifulSoup(GET.get(PAGE, headers={'User-Agent': 'Mozilla/5.0'}).text, 'html.parser').find_all('a', attrs={'class' : 'p-4 border border-gray-300 rounded'})

def CHECK_INTERNET_CONNECTION(self):
    INTERNET = False
    while not INTERNET:
        self.CONNECTION.setStyleSheet("""border-radius: 20px; background-color:qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:1 rgba(113, 205, 219, 255)); background-color: rgb(85, 85, 85); color:#37ff00; font: 75 13pt "MS Shell Dlg 2";""")
        self.CONNECTION.setText("CONNECTING.. 🌍")
        APP.processEvents()
        try:
            _ = GET.get('HTTPS://GOOGLE.COM', headers={'User-Agent': 'Mozilla/5.0'})
            self.CONNECTION.setText("CONNECTED ✅")
            self.CONNECTION.setStyleSheet("""border-radius: 20px; background-color:qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:1 rgba(113, 205, 219, 255)); background-color: rgb(85, 85, 85); color:#37ff00; font: 75 13pt "MS Shell Dlg 2";""")
            APP.processEvents()
            INTERNET = True
        except requests.RequestException as NO_INTERNET:
            self.CONNECTION.setText("NO INTERNET 🚫🌐")
            self.CONNECTION.setStyleSheet("""border-radius: 20px; background-color:qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:1 rgba(113, 205, 219, 255)); background-color: rgb(85, 85, 85); color:#ff1515; font: 75 13pt "MS Shell Dlg 2";""")
        time.sleep(1)

def DARK_MODE_FUNCTION(self):
    if self.centralwidget.styleSheet() == MAIN_STYLE_SHEET:
        self.centralwidget.setStyleSheet(DARK_MODE_STYLE_SHEET)
        self.B_DARK_MODE.setText("🌞")
    else:
        self.centralwidget.setStyleSheet(MAIN_STYLE_SHEET)
        self.B_DARK_MODE.setText("🌚")

def RUN_DATABASE(self):
    global DATABASE

    DATABASE = [CURSOR.execute("SELECT NAME, PHONE FROM {}".format(TABLE)).fetchall() for TABLE in ["BOTH", "MEN", "WOMEN"]]
    
    try:
        self.L_SUM_ALL_COUNT.setText(str(len(DATABASE[0]))) #########
        self.L_SUM_MEN_COUNT.setText(str(len(DATABASE[1]))) ######### SAME LABLES IN WINDOW 1 AND WINDOW 2
        self.L_SUM_WOMEN_COUNT.setText(str(len(DATABASE[2]))) #####
        
        self.L_DATABASE_ALL_COUNT.setText(str(len(DATABASE_START[0])))
        self.L_DATABASE_MEN_COUNT.setText(str(len(DATABASE_START[1])))
        self.L_DATABASE_WOMEN_COUNT.setText(str(len(DATABASE_START[2])))

        self.L_NEW_ALL_COUNT.setText(str(len(DATABASE[0]) - len(DATABASE_START[0])))
        self.L_NEW_MEN_COUNT.setText(str(len(DATABASE[1]) - len(DATABASE_START[1])))
        self.L_NEW_WOMEN_COUNT.setText(str(len(DATABASE[2]) - len(DATABASE_START[2])))
    except:
        pass

def WITCH_RADIO_BUTTON(self):
    if self.RADIO_ALL.isChecked():
        self.RADIO_BUTTON = "BOTH"
        self.URL = 'https://hub.23055.ir/search-official-expert?p_p_id=NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_pos=1&p_p_col_count=2&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_fname=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_lname=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_search=1&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_prvnc=23&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_cty=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_gnd=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_cnumber=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_fld=1217&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_delta=12&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_keywords=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_advancedSearch=false&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_andOperator=true&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_resetCur=false&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_cur='
    elif self.RADIO_MEN.isChecked():
        self.RADIO_BUTTON = "MEN" 
        self.URL = 'https://hub.23055.ir/search-official-expert?p_p_id=NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_pos=1&p_p_col_count=2&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_fname=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_lname=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_search=1&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_prvnc=23&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_cty=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_gnd=1&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_cnumber=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_fld=1217&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_delta=12&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_keywords=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_advancedSearch=false&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_andOperator=true&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_resetCur=false&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_cur='
    else:
        self.RADIO_BUTTON = "WOMEN"
        self.URL = 'https://hub.23055.ir/search-official-expert?p_p_id=NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_pos=1&p_p_col_count=2&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_fname=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_lname=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_search=1&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_prvnc=23&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_cty=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_gnd=2&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_cnumber=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_public_fld=1217&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_delta=12&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_keywords=&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_advancedSearch=false&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_andOperator=true&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_resetCur=false&_NetFormRecordsViewer_WAR_NetForm_INSTANCE_Bj31ZqhnvJ3m_cur='

def XLSX(NAMES, NUMBERS, RADIO_BUTTON):
    if NAMES != []:
        workbook = xlsxwriter.Workbook(RADIO_BUTTON + " - " + str(len(NAMES)) + '.XLSX')
        worksheet = workbook.add_worksheet()
        my_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        worksheet.set_column('A:A', 20, my_format), worksheet.set_column('B:B', 15, my_format)

        for ROW, (NAME, NUMBER) in enumerate(zip(NAMES, NUMBERS)):
            worksheet.write("A"+str(ROW + 1), NAME)
            worksheet.write("B"+str(ROW + 1), NUMBER)

        workbook.close()

def VCF(NAMES, NUMBERS, RADIO_BUTTON):
    if NAMES != []:
        NAMES = ["کارشناس " + ITEM for ITEM in NAMES]

        CONTACTS = [{'name': NAME, 'phone': NUMBER} for NAME, NUMBER in zip(NAMES, NUMBERS)]
        
        VCARD_LIST = []

        for CONTACT in CONTACTS:
            VCARD = vobject.vCard()
            VCARD.add('fn').value = CONTACT['name']
            TEL = VCARD.add('tel')
            TEL.value = CONTACT['phone']
            TEL.type_param = 'CELL'
            VCARD_LIST.append(VCARD)

        with open(RADIO_BUTTON + " - " + str(len(NAMES)) + '.VCF', 'w',encoding='utf-8') as VCF_FILE:
            for VCARD in VCARD_LIST:
                VCF_FILE.write(VCARD.serialize())

def ADD_CONTACTS(self, ID, NAME, PHONE): # ADD TO CONTACTS.DB
    WITCH_RADIO_BUTTON(self)

    CURSOR.execute("SELECT name, phone FROM %s WHERE name = ? AND phone = ?" % self.RADIO_BUTTON, (NAME, PHONE))
    EXISTING_CONTACT = CURSOR.fetchone()

    if EXISTING_CONTACT:
        return

    CURSOR.execute("INSERT INTO %s (ID, NAME, PHONE) VALUES (?, ?, ?)" % self.RADIO_BUTTON,  (ID, NAME, PHONE))
    
    CONNECT_SQLITE3.commit()
#---------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    APP = QApplication(sys.argv)
    MAIN_WINDOW = Main()
    WINDOW_2 = WINDOW_2()
    WIDGET = QtWidgets.QStackedWidget()
    WIDGET.setWindowFlags(QtCore.Qt.FramelessWindowHint)
    WIDGET.addWidget(MAIN_WINDOW)
    WIDGET.addWidget(WINDOW_2)
    WIDGET.setFixedHeight(880)
    WIDGET.setFixedWidth(560)
    WIDGET.show()
    APP.exec()
#---------------------------------------------------------------------------------------------------