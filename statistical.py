import os
import pandas as pd
import tkinter as tk
import tkinter.filedialog
import csv
import xlwt
from xlwt import *

# 定义一些常量
# originPath = 'data\\'  #数据所在文件夹
xuekeName = ['Agricultural Sciences', 'Biology & Biochemistry', 'Chemistry', 'Clinical Medicine', 'Computer Science',
             'Economics & Business', 'Engineering', 'Environment Ecology', \
             'Geosciences', 'Immunology', 'Materials Science', 'Mathematics', 'Microbiology', \
             'Molecular Biology & Genetics', 'Multidisciplinary', 'Neuroscience & Behavior', \
             'pharmacology & toxicology', 'Physics', 'Plant & Animal Science', 'Psychiatry Psychology', \
             'Social Sciences, General', 'Space Science']


def readExceptSch():
  with open(r'机构研究所名.txt') as file:
    return [i.split('\n')[0] for i in file.readlines()]


def readFileName(path):
  """
  获取指定目录下的所有文件名，即需要统计的数据所在的文件夹所有的xlsx文件
  :param path:
  :return:
  """
  fileName = []
  for root, dirs, files in os.walk(path):
    for file in files:
      if os.path.splitext(file)[1] == '.xlsx':
        fileName.append(os.path.join(file))
  return fileName


def readFile(path, originPath):
  """
  读取每一个文件，列重命名，并取得其中数据
  :param path: 文件路径
  :return: 筛选出的数据
  """
  column = ['range', 'schName', 'nation', 'p2', 'p3', 'p4', 'p5']

  data = pd.read_excel(originPath + path, skiprows=6, skipfooter=1, header=-1, names=column)
  # tempSchName = [i for i in data['schName'].values if i not in exceptSchName]
  data_from_chinaMainland = data[(data['nation'] == 'CHINA MAINLAND')]
  # data_from_chinaMainland = data_from_chinaMainland[data_from_chinaMainland('schName').isin(tempSchName)]
  return data_from_chinaMainland, len(data)


def getData(fileNameList, originPath):
  """
  获得所有的数据，存到列表
  :param fileNameList: 路径下需要统计的文件名列表
  :return: 所有读到的结果
  """
  result = []
  allLen = []
  for i in fileNameList:
    result.append(readFile(i, originPath)[0].reset_index(drop=True).drop(columns=['nation', 'p2', 'p3', 'p4', 'p5']))
    allLen.append(readFile(i, originPath)[1])
  print(result, allLen)
  return result, allLen


def mainTest(dataDir):
  # print(readFileName(r'data'))
  # exceptSchName = readExceptSch()
  allData = getData(readFileName(dataDir), dataDir)[0]
  length = getData(readFileName(dataDir), dataDir)[1]
  jigouData = allData[-1]
  jigouName = jigouData['schName']
  print(jigouData,len(jigouData))
  rangeData = []
  # for (i,j) in zip(xuekeName,allData[:-1]):
  #   print(j.rename(columns={'range':i},inplace=True))
  #   break

  final_dataFrame['schName'] = jigouName
  final_dataFrame['schRange'] = jigouData['range']
  # tempSchName = [i for i in final_dataFrame['schName'].values if i not in exceptSchName]
  # final_dataFrame = final_dataFrame[final_dataFrame['schName'].isin(tempSchName)]
  result = mergeData(allData[:-1])
  # print(allData[2])
  return result, length

def changeDataType(data):
  data['range']=[int(i) for i in data['range'] if i !=None]
  return data

def mergeData(data):
  t1 = pd.merge(final_dataFrame, data[0], how='left').rename(columns={'range': xuekeName[0]})
  t2 = pd.merge(t1, data[1], on='schName', how='left').rename(columns={'range': xuekeName[1]})
  t3 = pd.merge(t2, data[2], on='schName', how='left').rename(columns={'range': xuekeName[2]})
  t4 = pd.merge(t3, data[3], on='schName', how='left').rename(columns={'range': xuekeName[3]})
  t5 = pd.merge(t4, data[4], on='schName', how='left').rename(columns={'range': xuekeName[4]})
  t6 = pd.merge(t5, data[5], on='schName', how='left').rename(columns={'range': xuekeName[5]})
  t7 = pd.merge(t6, data[6], on='schName', how='left').rename(columns={'range': xuekeName[6]})
  t8 = pd.merge(t7, data[7], on='schName', how='left').rename(columns={'range': xuekeName[7]})
  t9 = pd.merge(t8, data[8], on='schName', how='left').rename(columns={'range': xuekeName[8]})
  t10 = pd.merge(t9, data[9], on='schName', how='left').rename(columns={'range': xuekeName[9]})
  t11 = pd.merge(t10, data[10], on='schName', how='left').rename(columns={'range': xuekeName[10]})
  t12 = pd.merge(t11, data[11], on='schName', how='left').rename(columns={'range': xuekeName[11]})
  t13 = pd.merge(t12, data[12], on='schName', how='left').rename(columns={'range': xuekeName[12]})
  t14 = pd.merge(t13, data[13], on='schName', how='left').rename(columns={'range': xuekeName[13]})
  t15 = pd.merge(t14, data[14], on='schName', how='left').rename(columns={'range': xuekeName[14]})
  t16 = pd.merge(t15, data[15], on='schName', how='left').rename(columns={'range': xuekeName[15]})
  t17 = pd.merge(t16, data[16], on='schName', how='left').rename(columns={'range': xuekeName[16]})
  t18 = pd.merge(t17, data[17], on='schName', how='left').rename(columns={'range': xuekeName[17]})
  t19 = pd.merge(t18, data[18], on='schName', how='left').rename(columns={'range': xuekeName[18]})
  t20 = pd.merge(t19, data[19], on='schName', how='left').rename(columns={'range': xuekeName[19]})
  t21 = pd.merge(t20, data[20], on='schName', how='left').rename(columns={'range': xuekeName[20]})
  t22 = pd.merge(t21, data[21], on='schName', how='left').rename(columns={'range': xuekeName[21]})

  return t22


def replaceSchName(data):
  # print(data['高校名称'].values)
  data_names = pd.read_excel(r'E-C_Name.xlsx')
  # chineseName = []
  # for (k, v) in enumerate(data_names['EnglishName']):
  #   for i in data['高校名称']:
  #     if i == v:
  #       chineseName.append(data_names.iloc[k].ChineseName)
  # print(data_names['ChineseName'])
  # for i in data['高校名称']:
  #   data_names[data_names['EnglishName']==i.values]
  #   for j in data['高校名称']:
  #     if k==j:
  #       chineseName.append(data_names.iloc[i]['ChineseName'])
  # print(chineseName,len(chineseName))
  # data['高校名称']=chineseName
  # return chineseName
  return [i.values[0] for i in [data_names[data_names['EnglishName'] == x].ChineseName for x in data['高校名称'].values]]

def set_style(name, height, bold=False):
  style = xlwt.XFStyle()  # 初始化样式

  font = xlwt.Font()  # 为样式创建字体
  font.name = name  # 'Times New Roman'
  font.bold = bold
  font.color_index = 4
  font.height = height

  # borders= xlwt.Borders()
  # borders.left= 6
  # borders.right= 6
  # borders.top= 6
  # borders.bottom= 6

  style.font = font
  # style.borders = borders

  return style

def format(lst,filePath):
  temp = readExceptSch()
  print(temp,len(temp))
  data1 = lst[0]
  allLen = [i + 1 for i in lst[1]]
  chinaLen = len(lst[0])
  print(chinaLen)
  data = data1.drop(index=data1[data1['schName'].isin(readExceptSch())].index, axis=0)

  data.rename(
    columns={'schName': '高校名称', 'schRange': '学校排名', 'Agricultural Sciences': '农业', 'Biology & Biochemistry': '生物学与生物科学',
             'Chemistry': '化学', 'Clinical Medicine': '临床医学', 'Computer Science': '计算机', 'Economics & Business': '经济与商业',
             'Engineering': '工程', \
             'Environment Ecology': '环境与生态', 'Geosciences': '地球', 'Immunology': '免疫', 'Materials Science': '材料',
             'Mathematics': '数学', 'Microbiology': '微生物', \
             'Molecular Biology & Genetics': '分子生物遗传', 'Multidisciplinary': '综合', 'Neuroscience & Behavior': '神经行为', \
             'pharmacology & toxicology': '药理毒理', 'Physics': '物理', 'Plant & Animal Science': '植物动物',
             'Psychiatry Psychology': '精神病心理', \
             'Social Sciences, General': '社会科学', 'Space Science': '空间科学'}, inplace=True)

  colSta, rowSta = counts(data)
  # print(colSta, rowSta,len(colSta),len(rowSta))
  data[u'数量'] = rowSta  # 每个学校 学科统计
  guojiData = readFile('z--jigou.xlsx', filePath)
  all = guojiData[1]
  print(guojiData[0],all)
  chineseNameList = replaceSchName(data)
  print(len(chineseNameList))
  data['高校名称'] = chineseNameList
  data.to_csv(r'result.csv', index=False, header=True, encoding='utf-8-sig',float_format='%.f')
  with open(r'result.csv', 'a+', newline='', encoding='utf-8-sig') as file:
    writer = csv.writer(file)
    writer.writerow(
      [u'进入排名中国高校', chinaLen - len(readExceptSch()), colSta[2], colSta[3], colSta[4], colSta[5], colSta[6], \
       colSta[7], colSta[8], colSta[9], colSta[10], colSta[11], colSta[12], colSta[13], colSta[14], colSta[15], \
       colSta[16], colSta[17], colSta[18], colSta[19], colSta[20], colSta[21], colSta[22], colSta[23], sum(colSta[2:])])
    writer.writerow([u'进入排名全球机构', all+1, allLen[0], allLen[1], allLen[2], allLen[3], allLen[4], allLen[5], allLen[6], \
                     allLen[7], allLen[8], allLen[9], allLen[10], allLen[11], allLen[12], allLen[13], allLen[14],
                     allLen[15], \
                     allLen[16], allLen[17], allLen[18], allLen[19], allLen[20], allLen[21], sum(allLen[:-1])])

  wb = Workbook(encoding='utf-8')
  sheet = wb.add_sheet(u'汇总表')
  # ----样式设置------
  style = xlwt.XFStyle()
  style.font.height = 180
  style1 = xlwt.XFStyle()
  style1.font.height = 200

  # row1 = sheet.row(0)
  # row1.height = 100 * 40
  col1 = sheet.col(0)
  col1.width = 380 * 20
  with open(r'result.csv', encoding='utf-8-sig') as f:
    read = csv.reader(f)
    l = 0
    for line in read:
      r = 0
      for i in line:
        if l == 0 :
          sheet.write(l, r, i, set_style('Arial', 200, True))
        elif l!=0 & r!=0:
          sheet.write(l, r, i, style=style)
        else:
          sheet.write(l, r, i, style=style)
          # sheet.col(0).width = 3000
        r = r + 1
      l = l + 1
    wb.save('汇总表.xls')

  return '统计完成，请退出！'




def counts(data):
  colSta = (len(data) - data.isna().sum()).values
  rowSta = (22 - data.isna().sum(axis=1)).values
  return colSta, rowSta




def chooseDir():
  filename = tkinter.filedialog.askdirectory()
  if filename != '':
    return filename
  else:
    return 0

def btn_click():
  chooseDir()

if __name__ == '__main__':
  window = tk.Tk()
  window.title(u'ESI统计分析软件')
  window.geometry('450x200')
  l = tk.Label(window, text='您好！使用本软件前，请仔细阅读使用说明。\n等待时间约为10s......',  width=80, height=2)
  l.pack()
  filePath = chooseDir() + "\\"
  # filePath = btn.mainloop()
  lab = tk.Label(window,text='\n您选择的路径是：\n'+filePath[:-1]+'\n',  width=100, height=2)
  lab.pack()
  # nameList = [i for i in readFileName(filePath)]
  # xueke = [i.split('.')[0] for i in readFileName(filePath)][:-1]
  # # print(xueke)
  final_dataFrame = pd.DataFrame(data=[], columns=xuekeName.append('schName'))
  resultData = mainTest(filePath)
  lb = tk.Label(window, text=format(resultData,filePath)+'\n\n如需再次统计请重启软件')
  lb.pack()


  window.mainloop()
  # readExceptSch(filePath)
  # makeWindows()
