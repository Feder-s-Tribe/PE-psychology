from PyQt5.QtWidgets import QMessageBox,QFileDialog

def show_error_message(statue,text,info=""):
    msg=QMessageBox()
    msg.setStyleSheet("QMessageBox{font-size:16px}")
    msg.setText(text)
    msg.setInformativeText(info)
    msg.setModal(False)
    if statue==1:
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle(u"错误")
    elif statue==2:
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle(u"警告")
    elif statue==3:
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle(u"问题")
    else:
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle(u"信息")
    msg.exec_()
    return 0

def get_path(style="file"):
    if style=="directory":
        path=QFileDialog.getExistingDirectory(None,u"选择目录","",QFileDialog.ShowDirsOnly)
    elif style=="file":
        path,_=QFileDialog.getOpenFileName(None,u"请选择一个文件","","Excel文件(*.xls *.xlsx)")
    return path