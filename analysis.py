from docxtpl import DocxTemplate
import pandas as pd

class analysis:
    data=pd.DataFrame
    lenrow=0
    col_name=[]
    def __init__(self,path) -> None:
        self.data=pd.read_excel(path,header=1)
        self.len=len(self.data)
        self.col_name=["name","gender","schoolID","nation","org","dep","sport","date","coachName","duration"]

    def generate(self):
        data=self.data#分析完成后的表
        word=DocxTemplate(r"sample\\SCL-90Scale.docx")
        for i in range(0,self.len):
            context=data.loc[i]