import tkinter as tk
from tkinter import filedialog
import pandas as pd
from opencc import OpenCC
from rich.progress import track
import docx

root = tk.Tk()
root.title('簡<->繁 轉換器')    #標題
root.geometry('+600+250')      #視窗位置
canvas1 = tk.Canvas(root,width=300,height=250)
canvas1.pack()

name = tk.Label(root,text='簡<->繁 轉換器')
name.config(font=20)
canvas1.create_window(150,60,window=name)

# 指定文件格式
class dif_type():
    def __init__(self,file):
        self.file = file

    def txt_type(self):
        txt_file = pd.read_table(self.file,header=None,encoding='utf-8')

        return txt_file

    def docx_type(self):
        word = docx.Document(self.file)  #讀取docx文件

        data = []
        for paragraph in word.paragraphs:   #將內容處存於list中
            data.append(paragraph.text)

        docx_file = pd.DataFrame()   #將list存于dataframe中
        docx_file[''] = pd.Series(data)

        return docx_file

# 獲取文件
def get_files():
    global convert_file
    file = filedialog.askopenfilename()

    final_type = dif_type(file)
    if file[-3:] == 'txt':
        convert_file = final_type.txt_type()
    if file[-4:] == 'docx':
        convert_file = final_type.docx_type()


get_button = tk.Button(text='選取文件',command=get_files,font=16)
canvas1.create_window(150,130,window=get_button)

# 文件格式(副檔名)
file_format = [('txt files','*.txt')]


def convert_complex():
    global convert_file
    data = []
    for i in track(range(convert_file.shape[0]),description='轉換中:'):
        open_cc = OpenCC('s2t.json')
        result = open_cc.convert(convert_file.loc[i][0])

        data.append(result)

    # 建立新的 dataframe
    new_file = pd.DataFrame()
    new_file['繁體文件：'] = pd.Series(data)

    # 保存文件
    import_file = filedialog.asksaveasfilename(filetypes=file_format)
    new_file.to_csv(import_file,line_terminator='\n\n',header=False,index=False)

convert_button = tk.Button(text='簡->繁',command=convert_complex,font=16)
canvas1.create_window(90,190,window=convert_button)

def convert_simple():
    global convert_file
    data = []
    for i in track(range(convert_file.shape[0]),description='轉換中:'):
        open_cc = OpenCC('t2s.json')
        result = open_cc.convert(convert_file.loc[i][0])
        data.append(result)

    # 建立新的 dataframe
    new_file = pd.DataFrame()
    new_file['簡體文件：'] = pd.Series(data)

    # 保存文件
    import_file = filedialog.asksaveasfilename(filetypes=file_format)
    new_file.to_csv(import_file,line_terminator='\n\n',header=False,index=False)

convert_button = tk.Button(text='繁->簡',command=convert_simple,font=12)
canvas1.create_window(210,190,window=convert_button)

root.mainloop()
