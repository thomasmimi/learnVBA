Sub 水平居中删除首行缩进()
'
' 表格大小调整 内容靠左 删除首行缩进 字体设置宋体五号 
' 段前0.5行，断后0行
' 行距1.25倍
' 表格首行加粗
' 
Dim t As Table
For Each t In ActiveDocument.Tables
	t.AutoFitBehavior (wdAutoFitContent) '根据内容调整表格大小
	't.AutoFitBehavior (wdAutoFitFixed)  '自动调整表格大小
	't.AutoFitBehavior (wdAutoFitWindow)  '根据窗口调整表格大小
    t.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '表格内容靠左
    Selection.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    t.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
    t.Range.ParagraphFormat.FirstLineIndent = 0   '取消首行缩进
    With Selection.Font
       .Size = 10.5
       .Name = "宋体"
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 2.5						'
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.25)
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
    Selection.Tables(1).Rows(1).Select
    Selection.Font.Bold = wdToggle
Next
End Sub
