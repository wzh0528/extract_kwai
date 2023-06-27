import pandas as pd
import os
import sys
from docx import Document

file_path = os.path.join("../", sys.argv[1])
print(file_path)
file_name = sys.argv[1].split('.')[0]


doc = Document(file_path)

name_list = []
id_list = []
video_list = []


for paragraph in doc.paragraphs[2:]:
    # print(paragraph.text)
    if '快手昵称' in paragraph.text:
        name_list.append(paragraph.text.split('：')[1])
    if '视频id' in paragraph.text:
        video_list.append(paragraph.text.split('：')[1])
    if '快手id' in paragraph.text:	
        id_list.append(paragraph.text.split('：')[1])

df = pd.DataFrame({'快手id':id_list, '快手昵称':name_list, '视频id':video_list})
writer = pd.ExcelWriter(f'../result_{file_name}.xlsx', engine='xlsxwriter')

df.to_excel(writer, index=False)
writer._save()

print("done")
