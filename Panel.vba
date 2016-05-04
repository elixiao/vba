    'Public train As New Collection '创建集合
    Dim cellsArray() As MyCell '创建保存MyCell的数组
    Public nameCol, noCol, upperCol, attrCol, fycCol As Integer '属性在表中的列数（重复利用）
    Dim maxRow As Integer '最大行数
    Dim index As Integer '数组位置计数器
    Dim dataSheet As String
    
    
Public Sub main()
    Debug.Print "进入main"
    dataSheet = "数据源"
    destiSheet = "架构图"
    nameCol = 4 '姓名列
    noCol = 3 '工号列
    upperCol = 6 '上级列
    bbCol = 7 '标保列
    fycCol = 8 'FYC列
    
    'attrCol = 7 '属性列
    'MsgBox Sheets(dataSheet).UsedRange.Rows.Count
    maxRow = Sheets(dataSheet).UsedRange.Rows.Count '最大行数
    Debug.Print "maxRow:" & maxRow
    ReDim cellsArray(maxRow) '定义数组的行数
    index = 0 '数组初始值
    
    
    '创建第一个主管节点（主管节点没有upper和right）
    Set cellsArray(0) = New MyCell '主管节点
    cellsArray(0).nameV = Sheets(dataSheet).Cells(2, nameCol).Value '主管姓名
    'cellsArray(0).attrV = Sheets(dataSheet).cells(2, attrCol).Value '主管职级
    cellsArray(0).noV = Sheets(dataSheet).Cells(2, noCol).Value '主管工号
    cellsArray(0).bbV = Sheets(dataSheet).Cells(2, bbCol).Value '标保
    cellsArray(0).rowV = 1 '主管行(不变)
    cellsArray(0).colV = 1 '主管列(不变)
    cellsArray(0).widthV = 0 '主管跨度(初值为0，实际值会改变)
    cellsArray(0).isLastV = 1 '是否最后(不变)
    cellsArray(0).childrenV = 0 '有几个子节点(初值为0，实际值会改变)
    cellsArray(0).fycV = Sheets(dataSheet).Cells(2, fycCol).Value 'FYC
    'MsgBox Sheets(dataSheet).Cells(2, fycCol).Value
    'cellsArray(0).descendantsV = maxRow - 2 '父节点子孙数
    Debug.Print cellsArray(0).upperV Is Nothing '测试上级节点是否为空节点
    'CellsArray(0).dayin
    'Set CellsArray(1) = New MyCell '注意，创建一个新对象的时候一定要用Set！
    'Set CellsArray(1) = CellsArray(0)
    Debug.Print "开始找孩子"
    findChildren cellsArray(0) '从主管节点开始不断寻找子节点
    'Dim i As Integer
    'ActiveSheet.Clear
    Sheets(destiSheet).Select
    
    Sheets(destiSheet).Cells.Clear
    drawFromZG
End Sub

'以主管为起点画架构图
Sub drawFromZG()
    Dim autoAttr As String
    For i = 0 To cellsArray(0).descendantsV
        'Debug.Print i
        cellsArray(i).dayin
        Application.DisplayAlerts = False '取消合并单元格的提示
        
'If False Then
        '自动发现节点的属性
        If cellsArray(i).upperV Is Nothing Then
            autoAttr = "(主管1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 39
        ElseIf cellsArray(i).descendantsV >= 4 Then '高级中支主任
            autoAttr = "(1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 40
            
        ElseIf cellsArray(i).descendantsV >= 2 Then '中支主任
            autoAttr = "(1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 6
        ElseIf cellsArray(i).descendantsV >= 1 Then '准中支主任
            autoAttr = "(1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 20
        Else
            autoAttr = "" '业务员
        End If
'End If
        
        
        If CDbl(cellsArray(i).fycV) >= 3000 Then
        
            With Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV))
            .Font.ColorIndex = 3
            '.Interior.ColorIndex = 40
            '.Interior.Pattern = xlPatternGray8
            '.Interior.PatternColorIndex = 36
            
            End With
        
        End If
        
        
        
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Merge '合并单元格
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0)).Value = cellsArray(i).nameV & autoAttr & Chr(10) & Round(CDbl(cellsArray(i).fycV), 0)
        
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).HorizontalAlignment = xlCenter '文字居中
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Borders.LineStyle = xlContinuous '添加边框
        
    Next i
End Sub

'以部经理为起点画架构图
Sub drawFromBJL()
    Dim autoAttr As String
    For i = 0 To cellsArray(0).descendantsV
        'Debug.Print i
        cellsArray(i).dayin
        Application.DisplayAlerts = False '取消合并单元格的提示
        '自动发现节点的属性
        If cellsArray(i).upperV Is Nothing Then
            autoAttr = "(部经理1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 33
        ElseIf cellsArray(i).upperV.noV = cellsArray(0).noV Then
            autoAttr = "(主管1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 39
        ElseIf cellsArray(i).descendantsV >= 4 Then '高级中支主任
            autoAttr = "(1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 40
            
        ElseIf cellsArray(i).descendantsV >= 2 Then '中支主任
            autoAttr = "(1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 20
        ElseIf cellsArray(i).descendantsV >= 1 Then '准中支主任
            autoAttr = "(1+" & cellsArray(i).descendantsV & ")"
            Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Interior.ColorIndex = 6
        Else
            autoAttr = "(1+" & cellsArray(i).descendantsV & ")" '业务员
        End If
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Merge '合并单元格
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0)).Value = cellsArray(i).nameV & Chr(10) & autoAttr & Chr(10) & cellsArray(i).fycV
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).HorizontalAlignment = xlCenter '文字居中
        Range(Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0), Cells(cellsArray(i).rowV + 0, cellsArray(i).colV + 0 + cellsArray(i).widthV)).Borders.LineStyle = xlContinuous '添加边框
        
    Next i

End Sub


'寻找指定节点的子节点
Sub findChildren(father As MyCell)
    Dim childrenCount As Integer '直增数量计数器
    childrenCount = 0 '直增初始值为0（没有直增）
    
    Dim former As New MyCell '定义上一个节点，用于指向右边兄弟节点
    Set former = father '上一个节点的初始值就是父节点本身
    Debug.Print "baba"
    '为什么从3开始？因为第一行是标题行，第二行是根节点行
    For i = 3 To maxRow
        If Sheets(dataSheet).Cells(i, upperCol).Value = father.noV Then '如果父节点工号是给定根节点工号（表明发现直增）
            
            childrenCount = childrenCount + 1 '直增加1
            index = index + 1 '数组位置计数器加1
            Set cellsArray(index) = New MyCell '新建节点（在新的数组位置上添加新节点）
            
            '设置新节点的属性(行数根据父节点来，列数根据左节点来)
            cellsArray(index).rowV = father.rowV + 1 '子节点行数=父节点行数+1
            cellsArray(index).upperV = father '指定该节点的父节点
            cellsArray(index).widthV = 0 '新节点的初始跨度都为0，跨度只有在insert的时候才会改变
            cellsArray(index).nameV = Sheets(dataSheet).Cells(i, nameCol) '设置姓名
            cellsArray(index).fycV = Sheets(dataSheet).Cells(i, fycCol) '设置fyc
           ' MsgBox "i=" & i & ",nameCol=" & nameCol & ",fycCol=" & fycCol & ",nameV=" & Sheets(dataSheet).Cells(i, fycCol) & ",fycV=" & Sheets(dataSheet).Cells(i, fycCol)
            'cellsArray(index).attrV = Sheets(dataSheet).cells(i, attrCol) '设置属性
            cellsArray(index).noV = Sheets(dataSheet).Cells(i, noCol) '设置工号
            
            Debug.Print "目前新节点是：" + cellsArray(index).nameV
            If Not former Is father Then '如果上一个节点不是父节点，说明former是它左边的兄弟节点
                former.rightV = cellsArray(index) '设置former右边节点为新节点
                former.isLastV = 0 '设置former的isLast的值为0，表明former不是father节点的最后一个子节点
                Debug.Print former.nameV
                cellsArray(index).colV = former.colV + former.widthV + 1 '列数等于former节点列数+former节点的跨度
                Set former = cellsArray(index) '将上一个节点更换为当前节点
                Debug.Print father.nameV
                father.insert '父节点要插入一行(父节点的单元格要右移）
                Debug.Print "wawa"
            Else '说明当前节点是father的第一个节点
                cellsArray(index).colV = father.colV '第一个节点的列数肯定是1（大错特错！）第一个节点列数肯定是父节点列数
                Set former = cellsArray(index) '将former节点设置为father的第一个子节点
            End If
            
            '一旦找到子节点，马上让子节点寻找自己的子节点（递归寻找，一步到位），关键是在这个时候不知道自己是不是最右节点
            '不妨假设不是最右节点，没有影响的，因为哈哈，我设计好了，列数=former+width+1
            cellsArray(index).isLastV = 0
            
            If childrenCount = 1 Then '如果子节点计数器是1，表明当前节点是father节点的第一个子节点
                Debug.Print "进入down"
                father.downV = cellsArray(index)
                Debug.Print "出去down"
            End If
            
            findChildren cellsArray(index)
           
        End If
    Next i
    
    If Not former Is father Then '用这种方式比较两个对象
        Debug.Print "hello"
        former.isLastV = 1 '最后一个former肯定是father节点的最后一个子节点
    End If
    
    father.childrenV = childrenCount '指定父节点的孩子数量
    
    Dim countDescendants As Integer
    countDescendants = 0
    If childrenCount = 0 Then
        father.descendantsV = 0
    Else
        Dim getChild As MyCell
        Set getChild = father.downV
        Do While Not getChild Is Nothing
            countDescendants = countDescendants + getChild.descendantsV + 1
            Set getChild = getChild.rightV
        Loop
    End If
    father.descendantsV = countDescendants
End Sub
