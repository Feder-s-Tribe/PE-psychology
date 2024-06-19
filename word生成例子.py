import os
import pandas as pd
from docxtpl import DocxTemplate,InlineImage
from docx.shared import Mm

data=pd.read_excel(r'耕地转包.xlsx',header=1)
lenrow=len(data)
col_name=data.columns
data=data.fillna('')

word=DocxTemplate(r'耕地转包调查表-40张.docx')

print(col_name)

for i in range(lenrow):
	context=data.loc[i]
	contexts=dict(zip(col_name,context))
	image_path=r'耕地转包照片\\'+str(context[0])+'.jpg'
	if os.path.exists(image_path):
		image=InlineImage(word,image_path,width=Mm(40))
		contexts["map"]=image
	word.render(contexts)
	word.save(r'耕地转包\\'+str(context[0])+'_耕地转包调查表.docx')
