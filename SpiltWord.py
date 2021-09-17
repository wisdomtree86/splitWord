from getCompanyCode import *
from win32com import client as wc
import os
import time
import shutil
def splitWord():
    _,companyname=getCompanyCodeData()
    word = wc.Dispatch("Word.Application")
    word.Visible=0
    word.DisplayAlerts=0
    for i in range(len(companyname)):

        #https://docs.microsoft.com/zh-cn/dotnet/api/
        print(companyname[i])
        isExists=os.path.exists(".\\year_report_2019\\%s"%companyname[i])
        if not isExists:
            os.makedirs(".\\year_report_2019\\%s"%companyname[i])
        directory = os.getcwd()
        document = word.Documents.Open(directory+"\\year_report_2019\\%s"%companyname[i]+"\\%s"%companyname[i]+".docx")
        activewindow=word.ActiveWindow
        actdocument = word.ActiveDocument
        word.Selection.EndKey(Unit=6)
        activewindow.ActivePane.View.Type=2
        #word.Selection.Range.Paragraphs.Style=actdocument.Styles(-2)
        activewindow.View.ShowHeading(1)
        if activewindow.View==5:
            activewindow.View=2
        else:
            activewindow.View=5
        word.Selection.HomeKey(Unit=6, Extend=1)
        activewindow.View.ShowAllHeadings()
        actdocument.Subdocuments.AddFromRange(word.Selection.Range)
        actdocument.Subdocuments.Expanded=1
        document.Close()
    word.Quit()


splitWord()
