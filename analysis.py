import os
from docxtpl import DocxTemplate
import pandas as pd

class analysis:
    def __init__(self,path) -> None:
        #Initialize the new data
        self.__data=pd.read_excel(path,usecols="H:CK")
        self.__newColums=self.__data.columns.values
        for i in range(len(self.__newColums)):
            self.__newColums[i]=self.__newColums[i][0:self.__newColums[i].index(u"、")]
        self.__data.columns=self.__newColums
        self.__gender={1:"男",2:"女"}
        self.__nation={
            1:"汉族",
            8:"彝族",
        }
        self.__data["2"]=self.__data["2"].replace(self.__gender,regex=True)
        self.__data["4"]=self.__data["4"].replace(self.__nation,regex=True)

        self.__db=pd.read_csv(r"sample\\db.csv",encoding="ANSI")
        self.len=len(self.__data)
        self.colNameInfo=["name","gender","schoolID","nation","org","dep","sport","date","coachName","duration","birthday"]
        self.colNameScore=["scoreBody","scoreForce","scoreRelation","scoreDep","scoreAnx","scoreHos","scoreHorr","scorePara","scoreSens","scoreOther"]

        self.__total={#original NO+11
            "scoreBody":["12","15","23","38","51","76","79"],
            "scoreForce":["14","21","39","49","56","62","66"],
            "scoreRelation":["16","17","32","45","47","48","52","63","73"],
            "scoreDep":["20","22","25","26","31","33","37","40","41","42","43","53","64","65"],
            "scoreAnx":["13","28","34","44","50","68"],
            "scoreHos":["35","70","82"],
            "scoreHorr":["24","36","57","58","61","71"],
            "scorePara":["19","29","54","69"],
            "scoreSens":["18","78","80","81"],
            "scoreOther":["27","30","46","55","59","60","67","72","74","75","77"]
        }

    def analysis(self):
        skip=0
        #Retrieve whether the name and date are in the database
        for row in self.__data.itertuples():
            name=row[1]
            date=row[8]       
            #if not, add data and calculate the result
            if self.__db.loc[(self.__db["name"]==name) & (self.__db["date"]==date)].empty:
                #Person infomation
                result=row[1:12]
                result+=tuple([0]*13) #total len of db list-11
                dbLen=len(self.__db.index)
                self.__db.loc[dbLen]=result

                #Change to dataframe
                dbData=pd.DataFrame([dict(zip(self.__newColums,row[1:]))])

                #Sub item score
                totalScore=0
                for sub in list(self.__total.keys()):
                    scorelist=self.__total.get(sub)
                    subScore=dbData.loc[0,scorelist].sum()
                    totalScore+=subScore
                    self.__db.loc[dbLen,sub]=subScore
                self.__db.loc[dbLen,"totalScore"]=totalScore
            else:
                skip+=1
                print(skip)
        self.__db.to_csv(r"sample\\db.csv",encoding="ANSI",index=False)


    def generate(self,savepath):
        word=DocxTemplate(r"sample\\SCL-90Scale.docx")
        for i in range(0,self.len):
            context=self.result.loc[i]
            contexts=dict(zip(self.colName,context))
            word.render(contexts)
            word.save(os.path.join(savepath,context[0]+str(context[9])+".docx"))

#test
if __name__ == '__main__':
    analysisResult=analysis(r"C:\\Users\\tengd\\Desktop\\272381692_按序号_运动员心理症状自评量表_25_25.xlsx")
    analysisResult.analysis()