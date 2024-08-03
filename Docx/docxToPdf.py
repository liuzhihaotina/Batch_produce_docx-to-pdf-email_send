# 如果该脚本运行失败，请尝试dtop.py  ----2024/8/2
# 注意：程序文件名不能为docx2pdf.py
# 注意：运行前要先保存
# 注意：运行convert函数前要先打开 Word
from docx2pdf import convert
import os
# # import glob
from pathlib import Path
import pythoncom

path = os.getcwd() + '/'
p = Path(path) # 初始化构造Path对象
file_list = list(p.glob("results/docx/动力/*.docx"))
print(file_list)
# print(len(file_list))

# 测试
# print(os.getcwd())
# convert('./template.docx', './template.pdf')
# 测试发现默认输出文件放在 文档 文件夹中
# os.chdir()
# exit()
for i in file_list:
    # 哈尔滨工业大学（深圳）推免面试成绩单-姓名
    # print(str(i))
    stu_name = str(i).split('-')[-1].split('.')[0]
    print(stu_name)
    pdf_name = f"哈尔滨工业大学（深圳）推免面试成绩单-{stu_name}.pdf"
    print(pdf_name)
    # major = str(i).split('\\')[-1].split('-')[0]
    # print(major)
    # 需要手动更改学科名称
    major = '动力'

    # output_path = './results/pdf/'+major+'/'+pdf_name
    # pdf的绝对路径
    output_path ='D:/控制面试/Docx - 20230727/Docx/results/pdf/'+major+'/'+pdf_name
    print(output_path)
    convert(i, output_path, keep_active=True)

    