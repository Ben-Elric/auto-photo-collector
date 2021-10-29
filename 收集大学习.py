import warnings
from win32com.client import Dispatch
import requests
import os
import zipfile

students = {}

def DownloadFile(name,url):
    try:
        print("[Info] 正在下载：" + name)
        r = requests.get(url, timeout=60)
        with open("大学习截图收集/" + name + ".jpeg", "wb") as code:
            code.write(r.content)
    except requests.exceptions.RequestException as e:
        print(e)
        print("[Error] 下载失败！")
    except:
        print("[Error] 下载失败！")


def LoadStudentList():
    with open('./list.txt', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        for line in lines:
            students[line.rstrip('\r\n')] = False


def LoadWorkbook(path):
    app = Dispatch("Excel.Application") 

    workbook = app.Workbooks.Open(path)
    sheet = workbook.Sheets[0]

    fileIndex = 0
    nameIndex = 0
    for j in range(1,sheet.UsedRange.Columns.Count):
        if sheet.Cells(1,j).Value == "资源地址":
            fileIndex = j
        elif sheet.Cells(1,j).Value == "打卡身份":
            nameIndex = j
    
    for i in range(2,sheet.UsedRange.Rows.Count):
        try:
            url = sheet.Cells(i,fileIndex).Hyperlinks.Item(1).Address
            name = sheet.Cells(i,nameIndex).Value.rstrip('\r\n')
            if students[name] == True:
                continue
            DownloadFile(name,url)
            students[name] = True
        except:
            break
    
    workbook.Close()
    app.Quit()


def ZipFiles(filename):
    print("\n[Info] 正在压缩...")
    with zipfile.ZipFile(filename + '.zip','w') as target:
        for i in os.walk('大学习截图收集'):
            for n in i[2]:
                target.write(''.join((i[0],'\\',n))) 
    DeleteImages()
    print("[Info] 压缩已完成")


def CheckStudents():
    print("\n未打卡名单：", end='')
    first = True
    for name,value in students.items():
        if value == False:
            if not first:
                print("，", end='')
            else:
                first = False
            print(name, end='')

def remove_dir(dir):
    dir = dir.replace('\\', '/')
    if(os.path.isdir(dir)):
        for p in os.listdir(dir):
            remove_dir(os.path.join(dir,p))
        if(os.path.exists(dir)):
            os.rmdir(dir)
    else:
        if(os.path.exists(dir)):
            os.remove(dir)

def DeleteImages():
    try:
        remove_dir('大学习截图收集')
    except:
        pass
    
def DeleteZips():
    try:
        for files in os.listdir('.'):
            if files.endswith(".zip"):
                os.remove(os.path.join('.',files))
    except:
        pass


def Main():
    LoadStudentList()
    DeleteImages()
    DeleteZips()

    try:
        os.mkdir('大学习截图收集')
    except:
        pass
    
    dataPath = ''
    for root,dirs,files in os.walk("."):
            for file in files:
                filename = os.path.abspath(file)
                if "大学习"  in filename and not "~" in filename:
                    dataPath = filename
                    break

    if(dataPath == ''):
        print('[FATAL] 未找到大学习数据文件！')
        print('[FATAL] 请将相关数据文件放到当前目录下')
    else:
        LoadWorkbook(dataPath)
        ZipFiles(dataPath.rstrip('.xlsx'))
        CheckStudents()

Main()
input()