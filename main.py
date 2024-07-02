import ctypes,sys
from PyQt5.QtWidgets import QApplication,QMainWindow

#ui
from _ui.main_ui import *
from _ui.ui_function import *

#function
from analysis import analysis

myappid="PE_Psychology"
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

#Bound signal of main UI
class main_ui(Ui_MainWindow):
    def __init__(self,MainWindow):
        super().__init__()
        self.setupUi(MainWindow)
        self.initUI()

    #QLineEdit set path
    def Q_path(self,result,style="file"):
        path=get_path(style)
        result.setText(path)

    #Save result
    def save_result(self):
        #Check the input
        if self.lineEdit_input.text()=="" and self.checkBox_input.isChecked():
            show_error_message(1,u"请选择文件")
            return 0
        elif self.lineEdit_save.text()=="" and self.checkBox_output.isChecked():
            show_error_message(1,u"请选择保存路径")
            return 0
        elif not (self.checkBox_input.isChecked() or self.checkBox_output.isChecked()):
            show_error_message(1,u"请勾选工作内容")
            return 0
        
        #Run function
        result="其中：\n"
        try:
            analysisResult=analysis(self.lineEdit_input.text())
            if self.checkBox_input.isChecked():
                skipAnaly=analysisResult.analysis()#Run analysis function
                result+=(str(skipAnaly)+u"条重复数据已跳过录入\n")
            if self.checkBox_output.isChecked():
                skipResult=analysisResult.generate(self.lineEdit_save.text())#generate forms of the results
                result+=(str(skipResult)+u"条重复结果已跳过生成\n")
        except Exception as e:
            show_error_message(1,u"生成失败")
            print(e)
        else:
            show_error_message(4,u"保存成功",result)

    #Initialize the UI function
    def initUI(self):
        self.pushButton_input.clicked.connect(lambda:self.Q_path(self.lineEdit_input))
        self.pushButton_save.clicked.connect(lambda:self.Q_path(self.lineEdit_save,"directory"))
        self.pushButton.clicked.connect(self.save_result)


#Main
if __name__ == '__main__':
    #Initialize UI
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = main_ui(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())