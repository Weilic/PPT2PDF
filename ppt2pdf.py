import os

import win32com.client


def PPT2PDF(file,pdfFiles):
    pdfName,ext = file.split('.')
    pdfName = pdfName + '.pdf'
    try:
        pdfFiles.index(pdfName)
    except ValueError:
        #print('OK')
        powerPoint = win32com.client.DispatchEx('Powerpoint.Application')
        file = os.getcwd() + '\\' + file
        pdfName = os.getcwd() + '\\' + pdfName
        nowFile = powerPoint.Presentations.Open(file)
        nowFile.SaveAs(pdfName, 32)
        print(pdfName+'导出成功！')
        powerPoint.Quit()
    else:
        pass


#testPath = 'E:/test'#测试目录
#dirFile = os.listdir(testPath)#测试目录文件
dirFile = os.listdir(os.getcwd())#获取脚本所在路径下的文件
#print(dirFile)
pptFiles = []
pdfFiles = []
for file in dirFile:
    if file.endswith('ppt') or file.endswith('pptx'):
        pptFiles.append(file)
    if file.endswith('pdf'):
        pdfFiles.append(file)
for pptfile in pptFiles:
    PPT2PDF(pptfile,pdfFiles)

