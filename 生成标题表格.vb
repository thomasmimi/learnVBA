Sub 生成标题表格()
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '遇到错误继续
    
    Dim mypath As String
    Dim myname As String
    Dim newpath As String
    Dim newname As String
    Dim row As Integer
    row = 2 '写入表格的起始行，从第二行开始写入
    '生成的表格和word文档，同名同目录
    mypath = ActiveDocument.Path
    myname = ActiveDocument.Name
    newname = Split(myname, ".doc")(0) '取得的文档名称去掉后缀.docx
    newpath = mypath & "\" & newname & ".xlsx" '表格保存目录
    Debug.Print newpath '打印当前文档路径
    '新建工作簿
    Dim exl As Object
    Set exl = CreateObject("excel.application") '创建excle对象
    Set wb = exl.Workbooks.Add '添加工作簿
    Set ws = wb.sheets(1) '
    ws.Name = "价格估算表" '修改工作表名称
          
    Dim myRange As Word.Range
    Dim num As String, content As String
    'Set ps = ActiveDocument.Bookmarks("\headinglevel").Range.Paragraphs '取得所有书签
    Set ps = ActiveDocument.Range.Paragraphs '取得所有书签

    For Each p In ps '对书签中每一个段落进行处理
        Set myRange = p.Range
        num = myRange.ListFormat.ListString '提取标题的序号
        If num <> "" Then '过滤不是标题的段落
            content = myRange.Text '提取标题名称
            ws.Cells(row, 1).Value = num '序号写入表格
            ws.Cells(row, 2).Value = content '标题内容写入表格
           row = row + 1 '下一行
        End If
    Next p
    '保存工作表并退出
    wb.SaveAs FileName:=newpath
    wb.Close True
    exl.Quit
    Set exl = Nothing
    Application.ScreenUpdating = True
End Sub