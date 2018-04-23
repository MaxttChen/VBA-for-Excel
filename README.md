# VBA-for-Excel
EXCEL的一些小程序;

***
1.EXCEL使用宏写VBA程序;
```
先调出开发工具;
再插入--选第一个--拉一个矩形框--点新建(在代码中生成"模块")
```


2.将一个文件夹下面的 2003格式的excel转换成2007格式
```
Sub btnTranst()

Dim strSourcePath, strTargetPath, strSourceFormat, strTargetFormat, strTraversalFileName As String
Dim openedWorkBook As Workbook


Application.ScreenUpdating = False '禁止屏幕更新，加快程序运行速度
Application.DisplayAlerts = False


strSourcePath = Worksheets(1).Range("B1") & "\"  '源路径
strSourceFormat = Worksheets(1).Range("B3")   '源格式
strTargetPath = Worksheets(1).Range("B6") & "\"  '目标路径
strTargetFormat = Worksheets(1).Range("B8") '目标格式

strSourceFormat = ".xls"   '源格式
strTargetFormat = ".xlsx" '目标格式

strTraversalFileName = Dir(strSourcePath & "*" & strSourceFormat)

Do While strTraversalFileName <> ""

Set openedWorkBook = Workbooks.Open(strSourcePath & strTraversalFileName, 0)

openedWorkBook.SaveAs Filename:=strTargetPath & Left(strTraversalFileName, Len(strTraversalFileName) - Len(strSourceFormat)) & strTargetFormat, FileFormat:= _
xlWorkbookDefault, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
, CreateBackup:=False

openedWorkBook.Close True

strTraversalFileName = Dir

Loop

Application.DisplayAlerts = True

Application.ScreenUpdating = True '恢复屏幕更新

MsgBox "转换完成"


End Sub

```
