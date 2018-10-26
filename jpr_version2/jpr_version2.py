import sys
import re
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QAction, QTableWidget, QTableWidgetItem, QVBoxLayout, \
QPushButton, QHBoxLayout, QGroupBox, QDialog, QFileDialog, QAction, QTabWidget, QLabel, QTableView, QAbstractItemView
from PyQt5.QtMultimedia import QMediaPlaylist, QMediaPlayer, QMediaContent
from PyQt5.QtGui import QIcon, QColor
from PyQt5.QtCore import pyqtSlot, QSize, QUrl, Qt
from shutil import copyfile
 
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
        self.tableRow = 6
        self.tableCol = 10
        self.row = 2
        self.words = []
        self.couple = []
        self.audiolist = []
        self.configFile = ".//configurationFile.xlsx"
        self.configFile_temp = ".//configurationFile_temp.xlsx"
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
            self.wb = openpyxl.load_workbook(fileName.strip())
            self.ws = self.wb.active
            self.fileReady = True
            self.statusbar.showMessage('succeed to load file...')
            self.row = 2
            self.tableWidget.clearContents()
        else:
            self.statusbar.showMessage('Fail to load file...')
        

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
 
        saveConfigButton = QPushButton('저장', self)
        saveConfigButton.clicked.connect(self.save_config)
        layout.addWidget(saveConfigButton)

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

        # Horizontal header
        self.logTable.setHorizontalHeaderItem(0, QTableWidgetItem("포장순번"))
        self.logTable.setHorizontalHeaderItem(1, QTableWidgetItem("거래처명"))
        self.logTable.setHorizontalHeaderItem(2, QTableWidgetItem("배송센터"))
        self.logTable.setHorizontalHeaderItem(3, QTableWidgetItem("박스"))
        self.logTable.setHorizontalHeaderItem(4, QTableWidgetItem("깔개"))
        self.logTable.setHorizontalHeaderItem(5, QTableWidgetItem("상부"))
        self.logTable.setHorizontalHeaderItem(6, QTableWidgetItem("하부"))
        self.logTable.setHorizontalHeaderItem(7, QTableWidgetItem("행잉"))
        self.logTable.setHorizontalHeaderItem(8, QTableWidgetItem("평대"))
        self.logTable.setHorizontalHeaderItem(9, QTableWidgetItem("배너"))

    def createConfigTable(self):
        self.configTable = QTableWidget()

        try:
            # load configurationFile
            cwb = load_workbook(self.configFile.strip())
        except:
            # message box
            print("Error")
            pass
        else:
            # Get matadata
            cws = cwb.active
            self.crow = cws.cell(row = 1, column = 1).value
            self.ccol = cws.cell(row = 1, column = 2).value

            # Configure table
            self.configTable.setRowCount(self.crow)
            self.configTable.setColumnCount(self.ccol)
            self.configTable.move(0,0)

            # Load data
            for i in range(self.crow):
                for j in range(self.ccol):
                    item = str(cws.cell(row = i+2, column = j+1).value).strip()
                    if item == 'None':
                        item = ''
                    self.configTable.setItem(i, j, QTableWidgetItem(item))

            for c in range(self.ccol):
                self.configTable.item(0, c).setBackground(QColor(255,255,0))

        self.configTable.itemChanged.connect(self.item_changed)


    def setTable(self):
        # shifting other rows
        for r in range(1, self.tableRow):
            for c in range(self.tableCol):
                try:
                    self.tableWidget.item(r,c).setBackground(QColor(255,255,255))
                    self.tableWidget.setItem(r-1 , c, self.tableWidget.item(r,c).clone())
                except:
                    pass

        # set current row
        self.tableWidget.setItem(self.tableRow - 1, 0, QTableWidgetItem(self.num))
        self.tableWidget.setItem(self.tableRow - 1, 1, QTableWidgetItem(self.partner))
        self.tableWidget.setItem(self.tableRow - 1, 2, QTableWidgetItem(self.parcel))
        self.tableWidget.setItem(self.tableRow - 1, 3, QTableWidgetItem(self.box))
        self.tableWidget.setItem(self.tableRow - 1, 4, QTableWidgetItem(self.kkar))
        self.tableWidget.setItem(self.tableRow - 1, 5, QTableWidgetItem(self.sang))
        self.tableWidget.setItem(self.tableRow - 1, 6, QTableWidgetItem(self.ha))
        self.tableWidget.setItem(self.tableRow - 1, 7, QTableWidgetItem(self.hang))
        self.tableWidget.setItem(self.tableRow - 1, 8, QTableWidgetItem(self.pyeng))
        self.tableWidget.setItem(self.tableRow - 1, 9, QTableWidgetItem(self.ban))

        for c in range(self.tableCol):
            self.tableWidget.item(self.tableRow - 1, c).setBackground(QColor(255,255,0))

    def read(self):
        if not self.fileReady:
            self.openFileNameDialog()

        #포장 순번
        self.num = str(self.ws.cell(row = self.row, column = 3).value[2:]).strip()
        print('<<<', self.num, '>>>', end = ' / ')
        self.couple.append('포장순번')
        self.parsing(self.num)

        #거래처명
        self.partner = str(self.ws.cell(row = self.row, column = 5).value).strip()
        print('거래처:', self.partner, end = ' / ')

        #택배발송
        self.parcel = str(self.ws.cell(row = self.row, column = 4).value).strip()
        print('배송센터:', self.parcel, end = ' / ')
        self.couple.append('배송센터')
        self.parsing(self.parcel)

        #박스
        self.box = str(self.ws.cell(row = self.row, column = 2).value).strip()
        print('박스:', self.box, end =' / ')
        self.couple.append('박스')
        self.parsing(self.box)


        #깔개
        self.kkar = str(self.ws.cell(row = self.row, column = 7).value).strip()
        print('깔개:', self.kkar, end =' / ')
        self.couple.append('깔개')
        self.parsing(self.kkar)


        #상부
        self.sang = str(self.ws.cell(row = self.row, column = 8).value).strip()
        print('상부:', self.sang, end =' / ')
        self.couple.append('상부')
        self.parsing(self.sang)


        #하부
        self.ha = str(self.ws.cell(row = self.row, column = 9).value).strip()
        print('하부:', self.ha, end =' / ')
        self.couple.append('하부')
        self.parsing(self.ha)


        #행잉
        self.hang = str(self.ws.cell(row = self.row, column = 10).value).strip()
        print('행잉:', self.hang, end =' / ')
        self.couple.append('행잉')
        self.parsing(self.hang)


        #평대
        self.pyeng = str(self.ws.cell(row = self.row, column = 11).value).strip()
        print('평대:', self.pyeng, end =' / ')
        self.couple.append('평대')
        self.parsing(self.pyeng)


        #배너
        self.ban = str(self.ws.cell(row = self.row, column = 12).value).strip()
        print('배너:', self.ban)
        self.couple.append('배너')
        self.parsing(self.ban)


        self.setTable()
        print(self.words)

    def parsing(self, val):
        if val == 'None' or val == '':
            self.couple.append('없음')

        else:
            if '(' in val or '/' in val:
                arr = re.split('[(/]',val)
                for i in range(len(arr)):
                    if ')' in arr[i]:
                        arr[i] = arr[i][:arr[i].index(')')]
                self.couple.append(arr)

            else:
                self.couple.append(val)

        self.words.append(self.couple)
        self.couple = [] 

    def load_audiolist(self):
        for item, value in self.words:
            if item == '배송센터':
                if value == '택배발송':
                    self.audiolist.append('택배발송')
            else:
                self.audiolist.append(item.strip())
                # 포장순번
                if item == '포장순번':
                    for i in range(3):
                        self.audiolist.append('_' + value[i])

                # in / or () case
                elif isinstance(value, list):
                    for name in value:
                        # '2개' case
                        if '2' in name:
                            self.audiolist.append('2')
                        else:
                            self.audiolist.append(name.strip())
                else:
                    self.audiolist.append(value.strip())

            # beep
            self.audiolist.append('beep')

    def speak(self):
        self.playlist.clear()
        for clip in self.audiolist:
            url = QUrl.fromLocalFile('.\\audio_clips\\' + clip + '.mp3')
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

        del self.words[:]
        del self.audiolist[:]
        self.read()
        self.load_audiolist()
        self.speak()


    @pyqtSlot()
    def cur_click(self):
        del self.words[:]
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

        del self.words[:]
        del self.audiolist[:]
        self.read()
        self.load_audiolist()
        self.speak()

    @pyqtSlot()
    def plus_row(self):
        self.crow = self.crow + 1
        self.configTable.setRowCount(self.crow)
        self.show()


    @pyqtSlot()
    def plus_col(self):
        self.ccol = self.ccol + 1
        self.configTable.setColumnCount(self.ccol)
        self.show()


    @pyqtSlot()
    def save_config(self):
        pass


    @pyqtSlot()
    def initialize(self):
        pass

    def item_changed(self,item):
        self.configTable.itemChanged.disconnect(self.item_changed)
        print(item.row(), item.column(), item.data(0))
        
        try:
            # load configurationFile
            cwb = load_workbook(self.configFile.strip())
        except:
            # message box
            print("configFile loading Error")
            pass
        else:
            cws = cwb.active
            previosData = str( cws.cell( row = item.row() + 2, column = item.column() + 1 ).value ).strip()
            print(previosData)

            fileName, _ = QFileDialog.getOpenFileName(self,"Open file", "", "All Files (*)")
            if fileName:
                print('ok')
                # copy file to local dir
                dst = "./audio_clips./" + str(item.row()) + '_' + str(item.column()) + '.wav'
                copyfile(fileName, dst)

                # change cell color
                self.configTable.item(item.row(), item.column()).setBackground(QColor(0,255,255))

                # update configFile
                cwb_w = Workbook(write_only=True)
                cws_w = cwb_w.create_sheet()

                cws_w.append([self.crow, self.ccol])

                for row in range(self.crow):
                    itemList = []
                    for col in range(self.ccol):
                        itemList.append(self.configTable.item(row, col).data(0))
                    cws_w.append(itemList)
                cwb_w.save(self.configFile_temp)

            else:
                self.configTable.setItem(item.row(), item.column(), QTableWidgetItem(previosData))
        self.configTable.itemChanged.connect(self.item_changed)



 
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())