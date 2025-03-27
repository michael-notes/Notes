# Notes
##  1. Snubber Circuit
### ROHM Snubber Circuit for Buck Converter IC 
https://fscdn.rohm.com/en/products/databook/applinote/ic/power/switching_regulator/buck_snubber_app-e.pdf
### TOSHIBA RC Snubbers for Step-Down Converters
https://toshiba.semicon-storage.com/info/application_note_en_20180901_AKX00078.pdf?did=63595
### Power Tips: Calculate an R-C Snubber in Seven Steps
https://www.ti.com/document-viewer/lit/html/SSZTBC7  
https://www.ti.com/lit/ta/ssztbc7/ssztbc7.pdf?ts=1723530928461&ref_url=https%253A%252F%252Fwww.google.com%252F
### RC Snubber Calculator Spreadsheet
https://paulorenato.com/index.php/electronics-diy/197-rc-snubber-calculator-spreadsheet
### 正确使用铝电解电容器的方法 CAT. No. C1001U
https://www.chemi-con.co.jp/cn/catalog/pdf/al-c/al-sepa-c/001-guide/al-technote-c-2020.pdf
![image](https://github.com/user-attachments/assets/ccb843df-946b-407f-a986-8424f3295e78)
### 铝电解电容器技术说明 CAT.8501
https://cn-nichicon.com/wp-content/uploads/technical-c.pdf
### Input Filter Design for Switching Power Supplies
#### Literature Number: SNVA538
https://www.ti.com/lit/an/snva538/snva538.pdf?ts=1737033376240
### Simple Success With Conducted EMI From DCDC Converters
#### Literature Number: SNVA489C
https://www.ti.com/lit/an/snva489c/snva489c.pdf?ts=1737030671475
### Frequency Response Analysis Tools for Push-pull Converter
https://u.dianyuan.com/bbs/u/37/1136215617.pdf
### ESR, Stability, and the LDO Regulator
https://www.ti.com/lit/an/slva115a/slva115a.pdf?ts=1743076299107
https://www.ti.com/cn/lit/an/zhca227/zhca227.pdf?ts=1743057562718
https://www.ti.com/document-viewer/lit/html/ssztbj1?keyMatch=LDO%20%E7%94%B5%E5%AE%B9%20ESR&tisearch=universal_search&f-technicalDocuments=Application%20note,Technical%20article




#工作表批量转换成工作簿
#VBA代码
Sub WorkbookToSheet()
     Application.DisplayAlerts = False
      Application.ScreenUpdating = False
      For i = 1 To ThisWorkbook.Sheets.Count
           ThisWorkbook.Sheets(i).Copy
          ActiveWorkbook.SaveAs ThisWorkbook.Path & "/" & ThisWorkbook.Sheets(i).Name, xlWorkbookDefault
           ActiveWorkbook.Close True
      Next
      Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "处理完成。", , "提醒"
End Sub

#代码解释与说明
Sub WorkbookToSheet()
    '关闭系统警告和消息提示,相同工作簿存在时会直接覆盖保存。
     Application.DisplayAlerts = False
    '关闭屏幕刷新，防止出现闪动
     Application.ScreenUpdating = False
     '遍历当前工作簿中的 第一个sheet 到 最后一个sheet 【ThisWorkbook.Sheets.Count=当前工作簿中的工作表的个数】
     For i = 1 To ThisWorkbook.Sheets.Count
        '复制当前工作簿中的第i个工作表
         ThisWorkbook.Sheets(i).Copy
        '工作表复制后，会成为活动工作薄,把活动的工作簿另存到当前工作簿的相同路径下，新的工作簿名字用被复制的工作表的名字，并采用默认Excel文件格式
         ActiveWorkbook.SaveAs ThisWorkbook.Path & "/" & ThisWorkbook.Sheets(i).Name, xlWorkbookDefault
        '关闭工作薄并保存
         ActiveWorkbook.Close True
    Next
    '前面强制关闭了屏幕刷新，程序结束前要恢复，否则会影响到平时的正常使用。
    Application.ScreenUpdating = True
    '前面强制关闭了警告和消息提示，程序结束前要恢复，否则会影响到平时的正常使用。
    Application.DisplayAlerts = True
    '提示已经处理完成。
    MsgBox "处理完成。", , "提醒"
End Sub
