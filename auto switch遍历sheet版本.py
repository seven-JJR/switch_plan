from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog,QHBoxLayout ,QWidget
from PyQt5.QtWidgets import QInputDialog,QFileDialog,QMessageBox
from autochange import Ui_MainWindow
import threading,openpyxl
work_loading_dict={}
issue_dict={}
class MainWindow(QMainWindow,Ui_MainWindow):
    mysignal = pyqtSignal(str) #代表信号以字符串方式发送, 只能定义在这里，不能定义在类外部或者init里面
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.Test_Case_Matrix_XC)
        self.pushButton_4.clicked.connect(self.original_XC)
        self.pushButton_5.clicked.connect(self.TCM_Plan_XC)
        self.mysignal.connect(self.prompt)

    def prompt(self):
        QMessageBox.information(MainWindow, '提示', 'Plan转换成功')

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
        sheets_names_Matrix=self.Matrix.sheetnames#获取每一个sheet的名称,返回['Iteration List', 'Application', 'BIOS', 'Multimedia', 'Mobile', 'Option', 'Network', 'UX']
        sheets_names_Matrix.pop(0)#删除第一个元素Iteration List
        self.testcase_sheet=self.TCM['test_case']
        self.i = 2  # 从excel TO TCM 的第二行开始加数据
        for every_sheet_name_Matrix in sheets_names_Matrix:
            every_sheet_Matrix=self.Matrix[every_sheet_name_Matrix]
            every_sheet_Matrix.maxrow=every_sheet_Matrix.max_row# 返回该sheet的最大行数
            every_sheet_Matrix_second_row=every_sheet_Matrix[2]#获取第二行的咨询，收集架构咨询
            every_sheet_Matrix_second_row_list=[]#获取第二行的咨询，收集架构咨询
            for every_cell in every_sheet_Matrix_second_row:#获取第二行的咨询，收集架构咨询
                every_sheet_Matrix_second_row_list.append(every_cell.value)#获取第二行的咨询，收集架构咨询
            for every_row in range(3,every_sheet_Matrix.maxrow+1):## 这里是匹配sheet的最大行数，除去前2行没有case，+1是因为range(2,9）行是循环2到8行.所以加1刚刚好
                every_row_list=[]
                for cell in every_sheet_Matrix[every_row]:# testplan_sheet[i] i是就代表testplan sheet的第几行，cell是代表这一行中的每一个单元格
                    every_row_list.append(cell.value)#cell.value每个单元格的值
                index=[index for index,item in enumerate(every_row_list) if item=='Incomplete']#巧用列表生成式和enumerate函数,找到每一行Incomplete的下标
                if len(index)==1: #代表只map了一个config
                    self.testcase_sheet.cell(column=1, row=self.i, value=every_row_list[0])
                    self.testcase_sheet.cell(column=2, row=self.i, value=every_sheet_name_Matrix)
                    self.testcase_sheet.cell(column=3, row=self.i, value=every_sheet_Matrix_second_row_list[index[0]]) # 因为index长度只有1，所以index[0]下标值就是对应第二行condfig的下标值
                    self.testcase_sheet.cell(column=4, row=self.i, value=every_sheet_Matrix_second_row_list[index[0]])
                    self.i=self.i+1
                else:#map了多个config，这时候就需要拆分了, index长度为多少就代表map了多少config, 就要分割成多少个列表
                    for every_index in index:#分割列表
                        self.testcase_sheet.cell(column=1, row=self.i, value=every_row_list[0])
                        self.testcase_sheet.cell(column=2, row=self.i, value=every_sheet_name_Matrix)
                        self.testcase_sheet.cell(column=3, row=self.i, value=every_sheet_Matrix_second_row_list[every_index]) # index下标值就是对应第二行condfig的下标值
                        self.testcase_sheet.cell(column=4, row=self.i, value=every_sheet_Matrix_second_row_list[every_index])
                        self.i = self.i + 1
        self.TCM.save(r'D:\items\TCM_plan_auto_switch\excel_to_TCMAPI.xlsx')#到这里TCM plan分割config部分就生成完成.
        self.TCM_result_plan= openpyxl.load_workbook(r'D:\items\TCM_plan_auto_switch\excel_to_TCMAPI.xlsx')  # 重新打开该excel根据origanl plan准备生成case结果
        self.original_plan=openpyxl.load_workbook(self.originaltestplan_path[0])#打开之前生成的original plan
        self.resultsheet=self.original_plan['result']
        self.resultsheet_maxrow=self.resultsheet.max_row
        result_all_list=[]#创建列表记录result 的所有case id用来计算哪些是测了的case
        for x in range(2,self.resultsheet_maxrow+1):#除去表头所以从2开始
            every_row_result_list=[]
            for every_cell_result in  self.resultsheet[x]:
                every_row_result_list.append(every_cell_result.value)
            result_all_list.append(every_row_result_list[0])#把每一行的第一个元素case ID传入到总的list里记录case iD
            work_loading_dict[every_row_result_list[0]]=every_row_result_list[2]#把case对应的loading增加到字典减脂对
            issue_dict[every_row_result_list[0]]=every_row_result_list[1]#把case对应的issue增加到字典减脂对
        self.testcasesheet_again=self.TCM_result_plan['test_case']
        self.testcasesheet_again_maxrow=self.testcasesheet_again.max_row
        i=1 #
        row_1=2#从第二行开始执行,眺过头标
        col_1=1#锁定第一列的ID值判断
        while i <= self.testcasesheet_again_maxrow: # 这里用while，用for循环会导致有的行被跳过执行，因为比如for x in rage（1，max)删除了第一行后继续执行x=2第二行，但是被删除了一行，原本的第二行变成第1眺过了
            if self.testcasesheet_again.cell(row=row_1,column=col_1).value not in result_all_list:#如果第一行第一列ID 不在result 列表里，代表这份case还没测完，删掉这一行.
                self.testcasesheet_again.delete_rows(row_1)
            else:#case iD有被找到，这是就要眺到下一行查找，如果没被找到被删了，就依然下次匹配第一行，因为原本第二行变成了第一行
                row_1=row_1+1
            i=i+1
        for jr in range(2,self.testcasesheet_again.max_row+1):#这里记得用新的最大行值，因为前面删除了很多行，所以变了
            list=[]
            for cells in self.testcasesheet_again[jr]:
                list.append(cells.value)
            self.testcasesheet_again.cell(row=jr,column=10,value=work_loading_dict[list[0]])# 根据case ID键找到对应的值
            self.testcasesheet_again.cell(row=jr, column=11, value=issue_dict[list[0]])#根据caee ID键找到对应的值
        self.TCM_result_plan.save((r'D:\items\TCM_plan_auto_switch\excel_to_TCMAPI.xlsx'))
        self.mysignal.emit('x')#发生字符串信号
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
