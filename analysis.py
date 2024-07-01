import os,io
from docxtpl import DocxTemplate,InlineImage
from docx.shared import Mm
from decimal import Decimal,ROUND_HALF_UP
# import plotly.graph_objects as go
import matplotlib.pyplot as plt
import pandas as pd

def rightRound(num,keep_n):
    if isinstance(num,float):
        num = str(num)
    return Decimal(num).quantize((Decimal('0.' + '0'*keep_n)),rounding=ROUND_HALF_UP)

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
        self.colNamePos=[x+"Pos" for x in self.colNameScore]
        self.colNameNeg=[x+"Neg" for x in self.colNameScore]
        
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
                result+=tuple([0]*33) #total len of db list-11
                dbLen=len(self.__db.index)
                self.__db.loc[dbLen]=result

                #Change to dataframe
                dbData=pd.DataFrame([dict(zip(self.__newColums,row[1:]))])

                #Sub item score
                totalScore=Positive=Negative=0
                for sub in list(self.__total.keys()):
                    scorelist=self.__total.get(sub)
                    subScore=dbData.loc[0,scorelist].sum()
                    subPos=(dbData[scorelist]>1).sum(axis=1)[0]
                    subNeg=(dbData[scorelist]==1).sum(axis=1)[0]
                    totalScore+=subScore
                    Positive+=subPos
                    Negative+=subNeg
                    self.__db.loc[dbLen,sub]=subScore
                    self.__db.loc[dbLen,sub+"Pos"]=subPos
                    self.__db.loc[dbLen,sub+"Neg"]=subNeg
                self.__db.loc[dbLen,"totalScore"]=totalScore
                self.__db.loc[dbLen,"Positive"]=Positive
                self.__db.loc[dbLen,"Negative"]=Negative
            else:
                skip+=1
        self.__db.to_csv(r"sample\\db.csv",encoding="ANSI",index=False)

        #Return the number of existed record
        return skip


    def generate(self,savepath):
        #Initialize the Word sample and reread db.csv
        word=DocxTemplate(r"sample\\SCL-90Scale.docx")
        self.__db=pd.read_csv(r"sample\\db.csv",encoding="ANSI")
        skip=0
        colName=self.colNameInfo+self.colNamePos+self.colNameNeg+["totalScore","Positive","Negative"]

        #plt
        #Chinese label configation
        plt.rcParams['font.sans-serif']=['SimHei']
        plt.rcParams['axes.unicode_minus']=False
        plt.rcParams['figure.figsize']=(14.76,7.68)#Pixel/1000
        plt.rcParams['font.size']=20
        nameList=[u"躯体",u"强迫",u"人际",u"抑郁",u"焦虑",u"敌对",u"恐怖",u"偏执",u"精敏",u"认知",u"总均"]
        index=[x for x in range(len(nameList))]

        for i in range(0,self.len):
            #Read basic information from db.csv and combine it with the sample dict
            context=self.__db.loc[i,colName]
            name=context[0]
            fileName=name+str(context[7])+".docx"
            path=os.path.join(savepath,name)
            pathFile=os.path.join(path,fileName)

            #Check whether the result exist
            if os.path.exists(pathFile):
                skip+=1
                continue

            contexts=dict(zip(colName,context))

            #Draw the histogram image
            # traceBasic=[go.Bar(
            #     x=[u"躯体化",u"强迫症状",u"人际关系敏感",u"抑郁",u"焦虑",u"敌对",u"恐怖",u"偏执",u"精神病性",u"其他项目",u"总症状指数"],
            #     y=self.__db.loc[i,self.colNameScore+["totalScore"]]
            # )]
            # figureBasic=go.Figure(data=traceBasic)
            # # figureBasic.show()
            # figureBasic.write_image(r"D:\\a.jpg",format="jpeg",engine="kaleido")#引擎好像有问题

            plt.ylabel(u"分数")
            plt.ylim((0,10))
            plt.xticks(index,nameList)
            Score=self.__db.loc[i,self.colNameScore+["totalScore"]]
            Negative=self.__db.loc[i,self.colNameNeg+["Negative"]]
            count=[7,7,9,14,6,3,6,4,4,11,71]
            y=[0]*11
            for i in range(11):
                y[i]=rightRound(Score[i]/count[i],2)
            firstBar=plt.bar(range(len(y)),y)
            for data in firstBar:
                y=data.get_height()
                x=data.get_x()
                plt.text(x+0.15,y,str(y),va='bottom') # x+n为偏移值，可以自己调整，正好在柱形图顶部正中
            #Export binary image
            imageIO=io.BytesIO()
            plt.savefig(imageIO,dpi=300,format="png")
            plt.cla()
            
            contexts["histogramResult"]=InlineImage(word,imageIO,width=Mm(125),height=Mm(65))

            word.render(contexts)

            #Save Word
            if not os.path.isdir(path):
                os.mkdir(path)
            word.save(pathFile)

        return skip

#test
if __name__ == '__main__':
    analysisResult=analysis(r"C:\\Users\\tengd\\Desktop\\272381692_按序号_运动员心理症状自评量表_25_25.xlsx")
    skip=analysisResult.analysis()
    print(skip)
    skip=analysisResult.generate(r"C:\\Users\\tengd\\Desktop\\test")
    print(skip)