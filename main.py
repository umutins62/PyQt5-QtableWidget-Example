import pandas as pd
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QTableWidget, QApplication, QVBoxLayout, QLineEdit, QLabel, QHBoxLayout, \
    QTableWidgetItem, QMainWindow, QAction, QMenu, QFileDialog, QComboBox, QColorDialog, QGraphicsColorizeEffect, \
    QProgressBar
import sys
from PyQt5 import QtWidgets, QtCore, QtGui

from pynput.keyboard import Key, Controller
keyboard1 = Controller()


class excelac(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setUI()


    def contextMenuEvent(self, event):
        contextMenu = QMenu(self)
        openAct = contextMenu.addAction(self.cad6)
        quitAct = contextMenu.addAction(self.cad)
        borderAct = contextMenu.addAction(self.cad44)
        cmbAct = contextMenu.addAction(self.cad55)
        pprgAct = contextMenu.addAction(self.cad66)
        action = contextMenu.exec_(self.mapToGlobal(event.pos()))

    def setUI(self):

        self.setWindowTitle("PypalamutXL V 1.0.0")
        self.setGeometry(600, 300, 1110, 300)



        self.fx=QLineEdit()
        self.cmb=QComboBox()


        self.secilen=QLineEdit()
        self.formul=QLabel("Fx: ")






        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(1000)
        self.tableWidget.setRowCount(1000)

        self.tableWidget.horizontalHeader().setDefaultSectionSize(130)
        self.tableWidget.horizontalHeader().setSortIndicatorShown(True)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.horizontalHeader().setStyleSheet(":section {""background-color: WhiteSmoke ; }")
        self.tableWidget.verticalHeader().setStyleSheet(":section {""background-color: WhiteSmoke ; }")
        # self.tableWidget.setStyleSheet("selection-background-color: Lavender;")
        self.tableWidget.cellClicked.connect(self.getsamerowcell)
        self.tableWidget.itemSelectionChanged.connect(self.degis)
        self.tableWidget.itemChanged.connect(self.hucreislemi)

        self.headretitle()

        self.fx.textChanged.connect(self.yaz)


        v1box=QHBoxLayout()
        v1box.addWidget(self.secilen,5)
        v1box.addWidget(self.formul,5)
        v1box.addWidget(self.fx,90)




        h2box=QVBoxLayout()
        h2box.addLayout(v1box)
        h2box.addWidget(self.tableWidget)

        widget = QtWidgets.QWidget()
        widget.setLayout(h2box)
        self.toolbar_menu()
        self.setCentralWidget(widget)
        self.centralWidget().setLayout(h2box)

        self.show()

    def toolbar_menu(self):
        self.tb = self.addToolBar("Tool Bar")
        self.tb.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)


        self.yeni = QAction(QIcon("new.png"), "New", self)
        self.yeni.triggered.connect(self.yeni1)
        self.tb.addAction(self.yeni)

        self.cad1 = QAction(QIcon("open.png"), "Open", self)
        self.cad1.triggered.connect(self.dosya_ac)
        self.tb.addAction(self.cad1)

        self.cad5 = QAction(QIcon("save.png"), "Save", self)
        # self.cad5.triggered.connect(self.Diger_Hesapla)
        self.tb.addAction(self.cad5)

        self.cad6 = QAction(QIcon("delete.png"), "Clear", self)
        self.cad6.triggered.connect(self.sil)
        self.tb.addAction(self.cad6)

        self.cad = QAction(QIcon("align_right_32px.png"), "Text Alligment Center", self)
        self.cad.triggered.connect(self.celltextaligmant)

        self.cad44 = QAction(QIcon("color.png"), "Change Background Color", self)
        self.cad44.triggered.connect(self.border1)

        self.cad55 = QAction(QIcon("combo-box.png"), "Add ComboBox", self)
        self.cad55.triggered.connect(self.getwidget)

        self.cad66 = QAction(QIcon("progress-bar.png"), "Add Progress-bar", self)
        self.cad66.triggered.connect(self.getwidget1)

    def border1(self):
        color = QColorDialog.getColor().getRgb()
        print(color)
        selected = self.tableWidget.selectedItems()

        if selected:
            for item in selected:
                item.setBackground(QtGui.QColor(color[0],color[1],color[2],color[3]))

    def getwidget(self):

        selected = self.tableWidget.selectedItems()
        if selected:
            for item in selected:
                self.tableWidget.setCellWidget(item.row(),item.column(),self.cmb)

    def getwidget1(self):
        self.pbar = QProgressBar(self)
        selected = self.tableWidget.selectedItems()
        if selected:
            for item in selected:
                self.pbar.setValue(int(item.text()))
                self.tableWidget.setCellWidget(item.row(),item.column(),self.pbar)

    def headretitle(self):
        list = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
                "V", "W", "X", "Y", "Z"]
        labels = []
        for i in range(len(list)):
            labels.append(list[i])
        for i in range(len(list)):
            for j in range(len(list)):
                a = list[i] + list[j]
                labels.append(str(a))
        for i in range(len(list)):
            for j in range(len(list)):
                for t in range(len(list)):
                    a = list[i] + list[j] + list[t]
                    labels.append(str(a))
        self.tableWidget.setHorizontalHeaderLabels(labels)

    def yeni1(self):
        self.tableWidget.clear()

    def dosya_ac(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, "Dosya Aç",
                                                   "/Excel Şeç",
                                                   "Excel (*.xls *.xlsx *.xlsm)")


            self.df = pd.read_excel (str(fileName[0]))

            list=[]

            for col in self.df.columns:
                list.append(col)

            print(self.df)




            self.tableWidget.setColumnCount(len(list))

            self.tableWidget.setHorizontalHeaderLabels(list)

            self.tableWidget.setColumnCount(len(self.df.columns))
            self.tableWidget.setRowCount(len(self.df.index))

            for i in range(len(self.df.index)):
                for j in range(len(self.df.columns)):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(self.df.iat[i, j])))


            self.tableWidget.resizeColumnsToContents()
            self.tableWidget.resizeRowsToContents()

            self.tableWidget.setColumnCount(1000)
            self.tableWidget.setRowCount(1000)
            self.headretitle()
            self.statusBar().showMessage("Dosya başarı ile açıldı.")




        except:
            pass

    def toplama(self,a,b):
        top=a+b
        return top

    def Convert(self,str):
        self.list1=[]

        self.list1[:0]=str
        return self.list1

    def celltextaligmant(self):
        paths = []
        selected = self.tableWidget.selectedItems()
        if selected:
            for item in selected:
                if item.column()> 0:
                    paths.append(item.data(0))
                    item.setTextAlignment(4)

    def sil(self):
        try:
            paths = []
            selected = self.tableWidget.selectedItems()
            if selected:
                for item in selected:
                    if item.column()> 0:
                        paths.append(item.data(0))
                        item.setText("")
        except:
            print("Hata")

    def degis(self):
        try:
            self.getsamerowcell()
            if self.tableWidget.item(self.rownum, self.columnnume).text()!="":
                self.fx.setText(self.tableWidget.item(self.rownum, self.columnnume).text())
            else:
                self.fx.setText("")


        except:
            pass

    def yaz(self):

            try:
                self.getsamerowcell()
                self.tableWidget.setItem(self.rownum,self.columnnume,QTableWidgetItem(self.fx.text()))
            except:
                pass

    def cellheader(self):
        for i in range(self.tableWidget.rowCount()):

            rw = self.tableWidget.horizontalHeaderItem(i).text()

            print(rw)

    def getsamerowcell(self):

        try:
            self.rownum = self.tableWidget.currentRow()
            self.columnnume = self.tableWidget.currentColumn()
            coltext = self.tableWidget.horizontalHeaderItem(self.columnnume).text()
            self.secilen.setText(coltext+str(self.rownum+1))



        except :
            pass
    
    def hucreislemi(self):

            try:
                list1=[]

                self.getsamerowcell()
                str1=self.tableWidget.item(self.rownum, self.columnnume).text()
                list1[:0]=str1

                x=3

                if keyboard1.pressed(Key.enter):
                    for i in range(len(list1)):
                        if list1[0]=="=":

                            if list1[-1]==")":
                                for t in range(len(list1)):
                                    if list1[i]==",":
                                        a1=float(str1[9:i])
                                        b1=float(str1[i+1:len(list1)-1])
                                        c=str(self.toplama(a1,b1))
                                        self.tableWidget.setItem(self.rownum,self.columnnume,QTableWidgetItem(c))

                            elif list1[i]=="+":
                                a=float(str1[1:i].replace(",","."))
                                b=float(str1[i+1:len(list1)].replace(",","."))
                                self.tableWidget.setItem(self.rownum,self.columnnume,QTableWidgetItem(str(round(a+b,x))))
                            elif list1[i]=="-":
                                a=float(str1[1:i].replace(",","."))
                                b=float(str1[i+1:len(list1)].replace(",","."))
                                self.tableWidget.setItem(self.rownum,self.columnnume,QTableWidgetItem(str(round(a-b,x))))
                            elif list1[i]=="/":
                                a=float(str1[1:i].replace(",","."))
                                b=float(str1[i+1:len(list1)].replace(",","."))
                                self.tableWidget.setItem(self.rownum,self.columnnume,QTableWidgetItem(str(round(a/b,x))))
                            elif list1[i]=="*":
                                a=float(str1[1:i].replace(",","."))
                                b=float(str1[i+1:len(list1)].replace(",","."))
                                self.tableWidget.setItem(self.rownum,self.columnnume,QTableWidgetItem(str(round(a*b,x))))
                        else:
                            pass
            except:
                pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    pencere = excelac()
    sys.exit(app.exec())