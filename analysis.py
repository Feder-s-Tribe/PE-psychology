import os
from docxtpl import DocxTemplate
import pandas as pd

class analysis:
    def __init__(self,path) -> None:
        self.__data=pd.read_excel(path,header=1)
        self.len=len(self.data)
        self.colName=["name","gender","schoolID","nation","org","dep","sport","date","coachName","duration"]
        self.__scoreBody=["11","14","22","37","50","75","78"]
        self.__scoreForce=["13","20","38","48","55","61","65"]
        self.__scoreRelation=["15","16","31","44","46","47","51","62","72"]
        self.__scoreDep=["19","21","24","25","30","32","36","39","40","41","42","52","63","64"]
        self.__scoreAnx=["12","27","33","43","49","67"]
        self.__scoreHos=["34","69","81"]
        self.__scoreHorr=["23","35","56","57","60","70"]
        self.__scorePara=["18","28","53","68"]
        self.__scoreSens=["17","77","79","80"]
        self.__scoreOther=["26","29","45","54","58","59","67","71","73","74","76"]

    def analysis(self):
        #total score
        self.__data.insert(self.__data.shape[1],"score",0)
        self.colName.append("score")

        #后面加计算，并更新self.colName
        
        #Result that needs to be generate
        self.result=self.__data.loc[:,self.colName]

    def generate(self,savepath):
        word=DocxTemplate(r"sample\\SCL-90Scale.docx")
        for i in range(0,self.len):
            context=self.result.loc[i]
            contexts=dict(zip(self.colName,context))
            word.render(contexts)
            word.save(os.path.join(savepath,context[0]+str(context[9])+".docx"))