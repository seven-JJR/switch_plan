from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog,QHBoxLayout ,QWidget
from PyQt5.QtWidgets import QInputDialog,QFileDialog,QMessageBox
from autochange import Ui_MainWindow
import threading,openpyxl
class MainWindow(QMainWindow,Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.Test_Case_Matrix_XC)
        self.pushButton_4.clicked.connect(self.original_XC)
        self.pushButton_5.clicked.connect(self.TCM_Plan_XC)

    def Test_Case_Matrix_XC(self):
        t1=threading.Thread(target=self.Test_Case_Matrix)
        t1.setDaemon(True)
        t1.start()
    def original_XC(self):
        t2 = threading.Thread(target=self.Original_Test_Plan)
        t2.setDaemon(True)
        t2.start()
    def TCM_Plan_XC(self):
        t3 = threading.Thread(target=self.TCM_Plan)
        t3.setDaemon(True)
        t3.start()
    def Test_Case_Matrix(self):
        self.caseMatrix_path=QFileDialog.getOpenFileName(None,'选择TCM Test Case Matrix','*.xlsx')#返回('C:/Users/7/Downloads/RD各部門庫存掛賬統計與管控0409.xlsx', 'All Files (*)')
    def Original_Test_Plan(self):
        self.originaltestplan_path = QFileDialog.getOpenFileName(None, '选择Original Test Plan', '*.xlsx')  # 返回('C:/Users/7/Downloads/RD各部門庫存掛賬統計與管控0409.xlsx', 'All Files (*)')
    def TCM_Plan(self):
        self.Matrix = openpyxl.load_workbook(self.caseMatrix_path[0])#如上所示返回的是列表，第一个代表路径
        self.TCM = openpyxl.load_workbook(r'D:\items\TCM_plan_auto_switch\excel_to_TCMAPI.xlsx')#打开该excel
        self.testcase_sheet=self.TCM['test_case']
        self.i = 2  # 从excel TO TCM 的第二行开始加数据
        self.Applicationsheet=self.Matrix['Application']#
        self.BIOSsheet = self.Matrix['BIOS']
        self.Multimediasheet = self.Matrix['Multimedia']
        self.Mobilesheet = self.Matrix['Mobile']
        self.Optionsheet = self.Matrix['Option']
        self.Networksheet = self.Matrix['Network']
        self.UXsheet = self.Matrix['UX']
        self.Applicationsheet.maxrow=self.Applicationsheet.max_row# 返回该sheet的最大行数
        self.Applicationsheet_second_row=self.Applicationsheet[2]#获取第二行的咨询，收集架构咨询
        Applicationsheet_second_row_list=[]#获取第二行的咨询，收集架构咨询
        for every_cell in self.Applicationsheet_second_row:#获取第二行的咨询，收集架构咨询
            Applicationsheet_second_row_list.append(every_cell.value)#获取第二行的咨询，收集架构咨询
        for every_row in range(3,self.Applicationsheet.maxrow+1):## 这里是匹配Applicationsheet的最大行数，除去前2行没有case，+1是因为range(2,9）行是循环2到8行.所以加1刚刚好
            every_row_list=[]
            for cell in self.Applicationsheet[every_row]:# testplan_sheet[i] i是就代表testplan sheet的第几行，cell是代表这一行中的每一个单元格
                every_row_list.append(cell.value)#cell.value每个单元格的值
            index=[index for index,item in enumerate(every_row_list) if item=='Incomplete']#巧用列表生成式和enumerate函数,找到每一行Incomplete的下标
            if len(index)==1: #代表只map了一个config
                self.testcase_sheet.cell(column=1, row=self.i, value=every_row_list[0])
                self.testcase_sheet.cell(column=2, row=self.i, value='Application')
                self.testcase_sheet.cell(column=3, row=self.i, value=Applicationsheet_second_row_list[index[0]]) # 因为index长度只有1，所以index[0]下标值就是对应第二行condfig的下标值
                self.testcase_sheet.cell(column=4, row=self.i, value=Applicationsheet_second_row_list[index[0]])
                self.i=self.i+1
            else:#map了多个config，这时候就需要拆分了, index长度为多少就代表map了多少config, 就要分割成多少个列表
                for every_index in index:#分割列表
                    self.testcase_sheet.cell(column=1, row=self.i, value=every_row_list[0])
                    self.testcase_sheet.cell(column=2, row=self.i, value='Application')
                    self.testcase_sheet.cell(column=3, row=self.i, value=Applicationsheet_second_row_list[every_index]) # index下标值就是对应第二行condfig的下标值
                    self.testcase_sheet.cell(column=4, row=self.i, value=Applicationsheet_second_row_list[every_index])
                    self.i = self.i + 1
        self.TCM.save(r'D:\items\TCM_plan_auto_switch\excel_to_TCMAPI.xlsx')
from PyQt5.QtGui import QGuiApplication
from PyQt5.QtCore import Qt
import sys
if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)#，不管在开发的机器上窗口设置多大，用户机器使用的时候窗口根据用户设备尺寸大小自动调整窗口大小使其居中
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自适应分辨
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = MainWindow()
    MainWindow.show()
    sys.exit(app.exec())
