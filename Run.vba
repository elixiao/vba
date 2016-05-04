Sub 运行()
    'Range("a:a").insert
    'Columns(2).insert
    'Range(Cells(1, 3), Cells(1, 6)).Merge
    'MsgBox "hello"
    
    筛选
    Dim a As New Panel
    a.main
     'Range(cells(1, 1), cells(1, 5)).Merge '合并单元格
End Sub

Sub 筛选()
     Worksheets("人员关系表").Range("b1").AutoFilter field:=2, Criteria1:=Worksheets("数据源").Range("k1").Value

    '先清空上次遗留数据
     Worksheets("数据源").Range("a1:h100").Clear
     
     '复制筛选后的值
     Worksheets("人员关系表").Range("a:h").SpecialCells(xlCellTypeVisible).Cells.Copy
     
     '粘贴到新的表当中
     Worksheets("数据源").Range("a1").PasteSpecial Paste:=xlPasteValues
     Worksheets("数据源").Range("a1").PasteSpecial Paste:=xlPasteFormats
     
End Sub
