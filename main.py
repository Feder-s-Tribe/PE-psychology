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
        if self.lineEdit_input.text()=="":
            show_error_message(1,u"请选择文件")
            return 0
        if self.lineEdit_save.text()=="":
            show_error_message(1,u"请选择保存路径")
            return 0
        
        #Run function
        try:
            analysisResult=analysis(self.lineEdit_input.text())
            analysisResult.analysis()#Run analysis function
            analysisResult.generate(self.lineEdit_save.text())#generate forms of the results
            function
        except:
            show_error_message(1,u"生成失败")
        else:
            show_error_message(4,u"保存成功")

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