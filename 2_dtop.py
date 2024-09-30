# convert .doc or .docx files to .pdf files in bulk
# 该脚本使用WPS，确保电脑具备该软件
# 首先安装pywin32
import os
from win32com.client import Dispatch


def word2pdf(abs_path, obj_path, files):  # word到pdf的转换
    if os.path.exists(obj_path):
        print("目标文件夹已经存在")
    else:
        os.mkdir(obj_path)
        print("目标文件夹创建成功")
    file_count = 0
    # word = Dispatch("Word.Application") # Word
    word = Dispatch("kwps.Application")  # WPS->word
    for file in files:
        file_count += 1
        doc = word.Documents.Open(abs_path + file)
        stu_name = str(file).split('-')[-1].split('.')[0]
        print(stu_name)
        pdf_name = f"2025年哈工大（深圳）机电学院推免生预接收函-{stu_name}.pdf"
        print(pdf_name)
        if file.endswith(".doc"):
            doc.SaveAs(obj_path + file.replace(".doc", ".pdf"), FileFormat=17)
        else:
            doc.SaveAs(obj_path + pdf_name, FileFormat=17)
        doc.Close()
    word.Quit()
    return file_count


def get_file(abs_path):  # 获取待转化的word文件名
    filenames = []
    raw_names = os.listdir(abs_path)
    for raw_name in raw_names:
        if raw_name.endswith(".docx") or raw_name.endswith(".doc"):
            # raw_name1='D:/控制面试/Docx-20230727/Docx/results/docx/控制/'+raw_name
            filenames.append(raw_name)
            # stu_name = str(raw_name).split('-')[-1].split('.')[0]
            # print(stu_name)
            # pdf_name = f"哈尔滨工业大学（深圳）推免面试成绩单-{stu_name}.pdf"
            # print(pdf_name)
        else:
            continue
    return filenames


if __name__ == "__main__":
    dir_path = os.path.dirname(os.path.abspath(__file__))   #获取 绝对路径主目录
    major='控制'
    #路径斜杠左低右高/，否则也容易报错，必须是绝对路径，否则报错
    abs_path = f'{dir_path}/生成文档/word/{major}/'  # word文件存放地址
    obj_path = f'{dir_path}/生成文档/pdf/{major}/'  # 转换后的pdf文件存放地址
    files = get_file(abs_path)
    # print(len(files))
    # print(files[0])
    # print(type(files[0]))
    file_number = word2pdf(abs_path, obj_path, files)
    print("{}个word文件已转换为相应pdf文件".format(file_number))


