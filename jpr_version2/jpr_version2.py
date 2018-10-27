import sys
import re
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QAction, QTableWidget, QTableWidgetItem, QVBoxLayout, \
QPushButton, QHBoxLayout, QGroupBox, QDialog, QFileDialog, QAction, QTabWidget, QLabel, QTableView, QAbstractItemView, \
QMessageBox
from PyQt5.QtMultimedia import QMediaPlaylist, QMediaPlayer, QMediaContent
from PyQt5.QtGui import QIcon, QColor
from PyQt5.QtCore import pyqtSlot, QSize, QUrl, Qt
from shutil import copyfile
from os import listdir, remove
 
class App(QMainWindow):
 
    def __init__(self):
        super().__init__()
        self.player = QMediaPlayer()
        self.playlist = QMediaPlaylist()
        self.playlist.setPlaybackMode(QMediaPlaylist.Sequential)
        self.title = 'JPR Reader GUI version2'
        self.left = 50
        self.top = 100
        self.width = 1050
        self.height = 550
        self.fileReady = False
        self.tableRow = 5
        self.tableCol = 15
        self.row = 2
        self.audiolist = []
        self.configFile = ".//configurationFile.xlsx"
        self.dict = {'num': None, 'partner': None, 'parcel': None, 'exception': None, 'box': None}
        self.initUI()
 
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        
        # menu
        menubar = self.menuBar()
        filemenu = menubar.addMenu('File')
        fileAct = QAction('Open File', self)
        fileAct.setShortcut('Ctrl+O')
        filemenu.addAction(fileAct)
        fileAct.triggered.connect(self.openFileNameDialog)

        # status_bar
        self.statusbar = self.statusBar()
 
        # tabs
        self.tabs = QTabWidget()
        self.tab1 = QWidget(self)
        self.tab2 = QWidget(self)
        self.tabs.addTab(self.tab1,"조작")
        self.tabs.addTab(self.tab2,"설정")
        self.setCentralWidget(self.tabs)

        # tab1 gui
        self.createHorizontalButtons1()
        self.createLogTable()
        self.tab1.layout = QVBoxLayout()
        self.tab1.layout.addWidget(self.horizontalButtons1) 
        self.tab1.layout.addWidget(self.logTable) 
        self.tab1.setLayout(self.tab1.layout)

        # tab2 gui
        self.explanation = QLabel()
        self.explanation.setText(
            """<<< JPR reader 설정 테이블 >>>    
            이곳에서 JPR reader가 읽어주는 항목과 아이템을 설정할 수 있습니다.""")
        self.explanation.setAlignment(Qt.AlignCenter)
        self.createConfigTable()
        self.createHorizontalButtons2()
        self.tab2.layout = QVBoxLayout()
        self.tab2.layout.addWidget(self.explanation) 
        self.tab2.layout.addWidget(self.configTable)
        self.tab2.layout.addWidget(self.horizontalButtons2) 
        self.tab2.setLayout(self.tab2.layout)

 
        # Show widget
        self.show()

    def openFileNameDialog(self):    
        fileName, _ = QFileDialog.getOpenFileName(self,"Open file", "", "XML files (*.xml, *.xlsx)")
        if fileName:
            self.wb = load_workbook(fileName.strip())
            self.ws = self.wb.active
            self.fileReady = True
            self.row = 2
            self.statusbar.showMessage('succeed to load file...')
            self.logTable.clear()
        else:
            self.statusbar.showMessage('Fail to load file...')

        # set logTable's horizontal header
        self.logTable.setHorizontalHeaderItem(0, QTableWidgetItem("포장순번"))
        self.logTable.setHorizontalHeaderItem(1, QTableWidgetItem("거래처명"))
        self.logTable.setHorizontalHeaderItem(2, QTableWidgetItem("배송센터"))
        self.logTable.setHorizontalHeaderItem(3, QTableWidgetItem("특이사항"))
        self.logTable.setHorizontalHeaderItem(4, QTableWidgetItem("박스"))

        self.dict = {'num': None, 'partner': None, 'parcel': None, 'exception': None, 'box': None}
        col = 6
        i = 0
        header = str(self.ws.cell(row = 1, column = col + 1).value).strip()
        while header != 'None':
            print(header)
            self.dict[header] = None
            col = col + 1
            i = i + 1
            header = str(self.ws.cell(row = 1, column = col + 1).value).strip()

        self.tableCol = 5 + i
        print(self.tableCol)
        self.logTable.setColumnCount(self.tableCol)

        for col in range(5, self.tableCol):
            self.logTable.setHorizontalHeaderItem(col, QTableWidgetItem(list(self.dict.keys())[col]))
            
        

    def createHorizontalButtons1(self):
        self.horizontalButtons1 = QGroupBox("Controller")
        layout = QHBoxLayout()
 
        pre_button = QPushButton('', self)
        pre_button.clicked.connect(self.pre_click)
        pre_button.setIcon(QIcon('.\\img\\pre.png'))
        pre_button.setIconSize(QSize(250,250))
        layout.addWidget(pre_button) 
 
        cur_button = QPushButton('', self)
        cur_button.clicked.connect(self.cur_click)
        cur_button.setIcon(QIcon('.\\img\\cur.png'))
        cur_button.setIconSize(QSize(250,250))
        layout.addWidget(cur_button) 
 
        next_button = QPushButton('', self)
        next_button.clicked.connect(self.next_click)
        next_button.setIcon(QIcon('.\\img\\next.png'))
        next_button.setIconSize(QSize(250,250))
        layout.addWidget(next_button) 
 
        self.horizontalButtons1.setLayout(layout)

    def createHorizontalButtons2(self):
        self.horizontalButtons2 = QGroupBox("설정 변경")
        layout = QHBoxLayout()
 
        plusRowButton = QPushButton('행+', self)
        plusRowButton.clicked.connect(self.plus_row)
        layout.addWidget(plusRowButton) 
 
        plusColButton = QPushButton('열+', self)
        plusColButton.clicked.connect(self.plus_col)
        layout.addWidget(plusColButton) 

        init_button = QPushButton('초기화', self)
        init_button.clicked.connect(self.initialize)
        layout.addWidget(init_button) 
 
        self.horizontalButtons2.setLayout(layout)
 
    def createLogTable(self):
        # Create table
        self.logTable = QTableWidget()
        self.logTable.setRowCount(self.tableRow)
        self.logTable.setColumnCount(self.tableCol)
        self.logTable.move(0,0)

    def createConfigTable(self):
        self.configTable = QTableWidget()

        try:
            # load configurationFile
            cwb = load_workbook(self.configFile.strip())
        except:
            # message box
            string = "Can't load configFile."
            QMessageBox.question(self, 'Error', string, QMessageBox.Ok, QMessageBox.Ok)

        else:
            # Get matadata
            cws = cwb.active
            self.crow = cws.cell(row = 1, column = 1).value
            self.ccol = cws.cell(row = 1, column = 2).value

            # Configure table
            self.configTable.setRowCount(self.crow)
            self.configTable.setColumnCount(self.ccol)
            self.configTable.move(0,0)

            # Load data from configFile
            for i in range(self.crow):
                for j in range(self.ccol):
                    item = str(cws.cell(row = i+2, column = j+1).value).strip()
                    if item == 'None':
                        item = ''
                    self.configTable.setItem(i, j, QTableWidgetItem(item))

            # check if files exist
            arr = listdir('./audio_clips')
            for row in range(self.crow):
                for col in range(self.ccol):
                    # if 박스(0,0)
                    if row == 0 and col == 0:
                        continue;

                    # reset backgound color
                    self.configTable.item(row, col).setBackground(QColor(255,0,0))
                    
                    # if file exist, change background color
                    fname = str(row) + '_' + str(col) + '.mp3'
                    if fname in arr:
                        if row == 0:
                            self.configTable.item(row, col).setBackground(QColor(255,255,0))
                        else:
                            self.configTable.item(row, col).setBackground(QColor(0,255,255))

            # 박스(0,0)
            self.configTable.setItem(0, 0, QTableWidgetItem('박스'))
            self.configTable.item(0, 0).setBackground(QColor(255,255,255))
            
            # link the callback function            
            self.configTable.itemDoubleClicked.connect(self.item_doubleClicked)
            self.configTable.itemChanged.connect(self.item_changed)


    def setLogTable(self):
        # shifting other rows
        for r in range(1, self.tableRow):
            for c in range(self.tableCol):
                try:
                    self.logTable.item(r,c).setBackground(QColor(255,255,255))
                    self.logTable.setItem(r-1 , c, self.logTable.item(r,c).clone())
                except:
                    pass

        # set current row
        for idx, key in enumerate(list(self.dict.keys())):
            if type(self.dict[key]) is list:
                self.logTable.setItem(self.tableRow - 1, idx, QTableWidgetItem(' '.join(self.dict[key])))
                self.logTable.item(self.tableRow - 1, idx).setBackground(QColor(255,255,0))
            else:
                self.logTable.setItem(self.tableRow - 1, idx, QTableWidgetItem(self.dict[key]))
                self.logTable.item(self.tableRow - 1, idx).setBackground(QColor(255,255,0))

    def read(self):
        if not self.fileReady:
            self.openFileNameDialog()

        # 포장 순번
        self.dict['num'] = str(self.ws.cell(row = self.row, column = 3).value[2:]).strip()
        print('<<<', self.dict['num'], '>>>', end = ' / ')

        # 거래처명
        self.dict['partner'] = str(self.ws.cell(row = self.row, column = 5).value).strip()
        print('거래처:', self.dict['partner'], end = ' / ')

        # 배송센터
        self.dict['parcel'] = str(self.ws.cell(row = self.row, column = 4).value).strip()
        print('배송센터:', self.dict['parcel'], end = ' / ')

        # 특이사항
        self.dict['exception'] = str(self.ws.cell(row = self.row, column = 6).value).strip()
        print('특이사항:', self.dict['exception'], end = ' / ')

        # 박스
        self.dict['box'] = str(self.ws.cell(row = self.row, column = 2).value).strip()
        print('박스:', self.dict['box'], end = ' / ')

        # left things
        print(len(self.dict))
        for i in range(5, len(self.dict)):
            header = str(self.ws.cell(row = 1, column = i + 2).value).strip()
            self.dict[header] = str(self.ws.cell(row = self.row, column = i + 2).value).strip()
            print(header + ':', self.dict[header], end = ' / ')
            self.parsing(header, self.dict[header])

        print(self.dict)
        self.setLogTable()


    def parsing(self, key, val):
        if val == 'None' or val == '':
            self.dict[key] = None

        else:
            if '(' in val or '/' in val:
                arr = re.split('[(/]',val)
                for i in range(len(arr)):
                    if ')' in arr[i]:
                        arr[i] = arr[i][:arr[i].index(')')]
                    arr[i] = arr[i].strip()
                self.dict[key] = arr

            else:
                self.dict[key] = val

    def itemFromKeyVal(self, key, val):
        items = self.configTable.findItems(val, Qt.MatchExactly)
        if len(items) <= 0:
            # Error
            string = '('+key+', '+val+') '+'아이템을 찾을 수 없습니다.'
            QMessageBox.question(self, 'Error', string, QMessageBox.Ok, QMessageBox.Ok)
        else:
            for item in items:
                if self.configTable.item(0, item.column()).data(0) == key:
                    return item


    def load_audiolist(self):
        for key, val in self.dict.items():

            # 포장순번     
            if key == 'num':
                for i in range(len(self.dict['num'])):
                    self.audiolist.append('_' + val[i])
                self.audiolist.append('_번')

                # beep
                self.audiolist.append('_beep')

            # 택배발송
            elif key == 'parcel':
                if val == '택배발송':
                    self.audiolist.append('_택배발송')

                    # beep
                    self.audiolist.append('_beep')
            
            # 박스
            elif key == 'box':
                item = self.itemFromKeyVal('박스', val)
                if item:
                    self.audiolist.append(str(item.row()) + '_' + str(item.column()))

                    # beep
                    self.audiolist.append('_beep')

            elif key in ['partner', 'exception']:
                pass
            
            # general case
            else:
                # The case(val == None) will be ignored 
                if val == None:
                    pass
                
                # when val is list
                elif type(val) == list:
                    for idx, eachVal in enumerate(val):
                        item = self.itemFromKeyVal(key, eachVal)
                        if item:
                            if idx == 0:
                                self.audiolist.append('0_' + str(item.column())) # key
                            self.audiolist.append(str(item.row()) + '_' + str(item.column())) # val

                    # beep
                    self.audiolist.append('_beep')

                # when val is not list
                else:
                    item = self.itemFromKeyVal(key, val)
                    if item:
                        if val == '1' or key == val:
                            self.audiolist.append('0_' + str(item.column())) # key
                        else:
                            self.audiolist.append('0_' + str(item.column())) # key
                            self.audiolist.append(str(item.row()) + '_' + str(item.column())) # val

                        # beep
                        self.audiolist.append('_beep')

            

        print(self.audiolist)

    def speak(self):
        self.playlist.clear()
        for clip in self.audiolist:
            url = QUrl.fromLocalFile('./audio_clips/' + clip + '.mp3')
            #print(url)
            self.playlist.addMedia(QMediaContent(url))
        self.player.setPlaylist(self.playlist)
        self.player.play()
            

    @pyqtSlot()
    def pre_click(self):
        if not self.fileReady:
            pass
        
        else:
            if self.row == 2:
                self.statusbar.showMessage("Can't be previous.")
            else:
                self.row -= 1

        self.dict = self.dict.fromkeys(self.dict, None)
        del self.audiolist[:]
        self.read()
        self.load_audiolist()
        self.speak()


    @pyqtSlot()
    def cur_click(self):
        self.dict = self.dict.fromkeys(self.dict, None)
        del self.audiolist[:]
        self.read()
        self.load_audiolist()
        self.speak()


    @pyqtSlot()
    def next_click(self):
        if not self.fileReady:
            pass

        else:
            if self.row == self.ws.max_row:
                self.statusbar.showMessage("It's over.")
            else:
                self.row += 1

        self.dict = self.dict.fromkeys(self.dict, None)
        del self.audiolist[:]
        self.read()
        self.load_audiolist()
        self.speak()

#----------------------------------------------#

    @pyqtSlot()
    def plus_row(self):
        # change configTable
        self.crow = self.crow + 1
        self.configTable.setRowCount(self.crow)
        
        # fill the generated cell
        self.configTable.itemChanged.disconnect(self.item_changed)
        for col in range(self.ccol):
            self.configTable.setItem(self.crow - 1, col, QTableWidgetItem(''))
            self.configTable.item(self.crow - 1, col).setBackground(QColor(255,0,0))
        self.configTable.itemChanged.connect(self.item_changed)



    @pyqtSlot()
    def plus_col(self):
        # change configTable
        self.ccol = self.ccol + 1
        self.configTable.setColumnCount(self.ccol)

        # fill the generated cell
        self.configTable.itemChanged.disconnect(self.item_changed)
        for row in range(self.crow):
            self.configTable.setItem(row, self.ccol - 1, QTableWidgetItem(''))
            self.configTable.item(row, self.ccol - 1).setBackground(QColor(255,0,0))
        self.configTable.itemChanged.connect(self.item_changed)


    @pyqtSlot()
    def initialize(self):
        # remove configurable audio files
        arr = listdir('./audio_clips')
        p = re.compile('_.+')
        for file in arr:
            if p.match(file):
                continue
            remove('./audio_clips/' + file)

        # init configTable
        self.configTable.itemChanged.disconnect(self.item_changed) # lock

        self.configTable.clear()
        self.crow = 5
        self.ccol = 5
        self.configTable.setRowCount(self.crow)
        self.configTable.setColumnCount(self.ccol)

        # reset configTable item
        for row in range(self.crow):
            for col in range(self.ccol):
                self.configTable.setItem(row, col, QTableWidgetItem(''))
                self.configTable.item(row, col).setBackground(QColor(255,0,0))
        self.configTable.setItem(0, 0, QTableWidgetItem('박스'))
        self.configTable.item(0, 0).setBackground(QColor(255,255,255))

        self.configTable.itemChanged.connect(self.item_changed) # unlock

        # init configFile
        self.update_configFile()

    def item_doubleClicked(self,item):
        self.previousItem = item.data(0)
        print(self.previousItem)

    def update_configFile(self):
        cwb_w = Workbook(write_only=True)
        cws_w = cwb_w.create_sheet()

        cws_w.append([self.crow, self.ccol])

        for row in range(self.crow):
            itemList = []
            for col in range(self.ccol):
                itemList.append(self.configTable.item(row, col).data(0))
            cws_w.append(itemList)
        try:
            cwb_w.save(self.configFile)
        except:
                string = """Maybe the configFile is opened.
                            Try again after close the file."""
                QMessageBox.question(self, 'Error', string, QMessageBox.Ok, QMessageBox.Ok)

    def item_changed(self,item):
        self.configTable.itemChanged.disconnect(self.item_changed) # lock
        
        if item.row() == 0 and item.column() == 0:
            string = "이 항복은 바꿀 수 없습니다."
            QMessageBox.question(self, 'Error', string, QMessageBox.Ok, QMessageBox.Ok)
            self.configTable.setItem(item.row(), item.column(), QTableWidgetItem(self.previousItem))

        else:
            # get file name 
            fileName, _ = QFileDialog.getOpenFileName(self,"Open file", "", "All Files (*)")

            # count the number of key that has same name
            keys = self.configTable.findItems(item.data(0), Qt.MatchExactly)
            kcnt = 0
            for key in keys:
                if key.row() == 0:
                    kcnt = kcnt + 1

            # count the number of atribute that has same name
            atributes = self.configTable.findItems(item.data(0), Qt.MatchExactly)
            acnt = 0
            for atribute in atributes:
                if atribute.row() == 0:
                    pass
                elif atribute.column() == item.column():
                    acnt = acnt + 1

            # change is accepted only in case of uniqueness and existence and it is not 박스(0,0)
            if kcnt >= 2:
                string = "항목명이 같을 수 없습니다."
                QMessageBox.question(self, 'Error', string, QMessageBox.Ok, QMessageBox.Ok)
                self.configTable.setItem(item.row(), item.column(), QTableWidgetItem(self.previousItem))

            elif acnt >= 2:
                string = "같은 항목에 같은 이름의 아이템을 둘 수 없습니다."
                QMessageBox.question(self, 'Error', string, QMessageBox.Ok, QMessageBox.Ok)
                self.configTable.setItem(item.row(), item.column(), QTableWidgetItem(self.previousItem))

            elif fileName:
                # copy file to local dir
                dst = "./audio_clips./" + str(item.row()) + '_' + str(item.column()) + '.mp3'
                copyfile(fileName, dst)

                # change cell color
                if item.row() == 0:
                    self.configTable.item(item.row(), item.column()).setBackground(QColor(255,255,0))
                else:
                    self.configTable.item(item.row(), item.column()).setBackground(QColor(0,255,255))

                # update configFile
                self.update_configFile()
                
            else:
                string = "선택이 취소됨."
                QMessageBox.question(self, 'Error', string, QMessageBox.Ok, QMessageBox.Ok)
                self.configTable.setItem(item.row(), item.column(), QTableWidgetItem(self.previousItem))


        self.configTable.itemChanged.connect(self.item_changed) # unlock



 
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())