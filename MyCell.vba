
Private row$ '所在行号（例如2）
Private col$ '所在列号（例如2）
Private width$ '列跨度（例如4）,初始值为0
Private isLast$ '判断是否是每行最右边的
Private children$ '下面直增个数（例如5）
Private descendants$ '下面所有子节点个数

Private name As String '业务员的姓名（例如周亮）
Private attr As String '职级名称（例如中支主任）
Private no As String '工号
Private bb As String '标保
Private fyc As String 'fyc

Private upper As MyCell '上级表格（弄清上级关系）
Private right As MyCell '右边表格（弄清同级关系）
Private down As MyCell '下级表格（第一个子节点弄清下级关系）


'递归插入
Public Sub insert()
    If Not upper Is Nothing Then '如果上级不是空节点，说明不是根节点，还要递归调用upper.insert，否则只需要增加跨度
        Debug.Print attr
        Debug.Print "在insert开始"
        'Columns(col + 1).insert '插入一列
        width = width + 1 '跨度增加
        'moveRight '递归调用MoveRight函数，确保兄弟节点能够右移 moveright函数根本就没有作用嘛
        Debug.Print "在insert中间"
        'Application.DisplayAlerts = False '取消合并单元格的提示
        'Range(cells(row + 0, col + 0), cells(row + 0, col + 0 + width)).Merge '合并单元格
        Debug.Print "在insert的merge之后"
        upper.insert '调用父节点函数
        Debug.Print "在insert结束"
    Else
        Debug.Print "我进来啦"
        width = width + 1 '根节点跨度始终要增加的
        Debug.Print col + 0 + width
        'Range(cells(row + 0, col + 0), cells(row + 0, col + 0 + width)).Merge '合并单元格
    End If
End Sub

'插入后，右侧单元格全部右移
Public Sub moveRight()
    Debug.Print "进入moveRight模块"
    If isLast = 0 Then '只要不是最后一个就继续右移
        right.col = right.col + 1
        right.moveRight
    End If
    Debug.Print "出去了moveright"
End Sub

'打印自己
Public Sub dayin()
    Debug.Print "========================="
    Debug.Print "行：" + row
    Debug.Print "列：" + col
    Debug.Print "跨度：" + width
    Debug.Print "直增：" + children
    Debug.Print "子孙：" + descendants
    Debug.Print "姓名：" + name
    Debug.Print "职级：" + attr
    Debug.Print "工号：" + no
    
    If Not upper Is Nothing Then
        Debug.Print "推荐人：" + upper.nameV
    Else
        Debug.Print "没有推荐人"
    End If
    If Not right Is Nothing Then
        Debug.Print "右兄弟：" + right.nameV
    Else
        Debug.Print "我就是最右"
    End If
    If Not down Is Nothing Then
        Debug.Print "大弟子：" + down.nameV
    Else
        Debug.Print "没有增员"
    End If
    Debug.Print "========================="
End Sub


'获取行号值
Public Property Get rowV() As Integer
  rowV = row
End Property
'给行号赋值
Public Property Let rowV(ByVal setRow As Integer)
  row = setRow
End Property


'获取列号值
Public Property Get colV() As Integer
  colV = col
End Property
'给列号赋值
Public Property Let colV(ByVal setCol As Integer)
  col = setCol
End Property


'获取跨度值
Public Property Get widthV() As Integer
  widthV = width
End Property
'给跨度赋值（注意，要修改上级跨度值）
Public Property Let widthV(ByVal setWidth As Integer)
  width = setWidth
End Property


'获取是否最右值
Public Property Get isLastV() As Integer
  isLastV = isLast
End Property
'给是否最右赋值
Public Property Let isLastV(ByVal setisLast As Integer)
  isLast = setisLast
End Property


'获取直增值
Public Property Get childrenV() As Integer
  childrenV = children
End Property
'给直增赋值
Public Property Let childrenV(ByVal setChildren As Integer)
  children = setChildren
End Property


'获取后代值
Public Property Get descendantsV() As Integer
  descendantsV = descendants
End Property
'给后代赋值
Public Property Let descendantsV(ByVal setDescendants As Integer)
  descendants = setDescendants
End Property


'给姓名赋值
Public Property Let nameV(ByVal setName As String)
  name = setName
End Property
'获取姓名值
Public Property Get nameV() As String
  nameV = name
End Property

'给FYC赋值
Public Property Let fycV(ByVal setFyc As String)
  fyc = setFyc
End Property
'获取FYC值
Public Property Get fycV() As String
  fycV = fyc
End Property


'给属性赋值
Public Property Let attrV(ByVal setAttr As String)
  attr = setAttr
End Property
'获取属性值
Public Property Get attrV() As String
  attrV = attr
End Property


'给工号赋值
Public Property Let noV(ByVal setNo As String)
  no = setNo
End Property
'获取工号值
Public Property Get noV() As String
  noV = no
End Property


'给标保赋值
Public Property Let bbV(ByVal setBb As String)
  bb = setBb
End Property
'获取标保值
Public Property Get bbV() As String
  bbV = bb
End Property


'给上级赋值
Public Property Let upperV(ByRef setUpper As MyCell)
  Set upper = setUpper
End Property
'获取上级值
Public Property Get upperV() As MyCell
  Set upperV = upper
End Property


'给右级赋值
Public Property Let rightV(ByRef setRight As MyCell)
  Set right = setRight
End Property
'获取右级值
Public Property Get rightV() As MyCell
  Set rightV = right
End Property


'给下级赋值
Public Property Let downV(ByRef setDown As MyCell)
  Set down = setDown
End Property
'获取下级值
Public Property Get downV() As MyCell
  Set downV = down
End Property

