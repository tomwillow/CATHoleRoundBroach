Attribute VB_Name = "ListViewFace_Module"
Option Explicit

'原作者 百度知道：瑞安阿芳
'来自 VB ListView用5.0的后怎么显示网络线 http://zhidao.baidu.com/question/300375363.html

''''ListView （Common Control 5.0） 改善显示用''''''''''''''''
Private Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const LVS_EX_GRIDLINES                      As Long = &H1&
Public Const LVS_EX_CHECKBOXES                     As Long = &H4&
Public Const LVS_EX_FULLROWSELECT                  As Long = &H20&
Public Const LVM_FIRST                             As Long = &H1000
Public Const LVM_SETITEMSTATE                      As Long = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE                      As Long = (LVM_FIRST + 44)
Public Const LVIS_STATEIMAGEMASK                   As Long = &HF000
Public Const LVIF_STATE                            As Long = &H8
'***********************常数

Private Enum LISTVIEW_MESSAGES
    'LVM_FIRST = &H1000
    LVM_SETITEMCOUNT = (LVM_FIRST + 47)
    LVM_GETITEMRECT = (LVM_FIRST + 14)
    LVM_SCROLL = (LVM_FIRST + 20)
    LVM_GETTOPINDEX = (LVM_FIRST + 39)
    LVM_HITTEST = (LVM_FIRST + 18)
    LVM_DELETEALLITEMS = (LVM_FIRST + 9)
    LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
    LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
End Enum
'//设置ListView扩展格式------------------------------------------------------------
Public Sub SetExtendedStyle(lv As Object, ByVal lStyle As Long, ByVal lStyleNot As Long)
    Dim lNewStyle   As Long
    lNewStyle = SendMessageLong(lv.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lNewStyle = lNewStyle And Not lStyleNot
    lNewStyle = lNewStyle Or lStyle
    SendMessageLong lv.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lNewStyle
End Sub
''''''''''''''''''''''''''''''结束Listview'''''''''''''''''''''''''''''''''''''''''

'用法示例
'Private Sub Form_Load()
'  SetExtendedStyle Listview1, LVS_EX_GRIDLINES, 0
'  SetExtendedStyle Listview1, LVS_EX_FULLROWSELECT, 0
'End Sub

Public Sub InitListViewTeeth()
With Form1
    .ListViewTeeth.Appearance = ccFlat
    .ListViewTeeth.BorderStyle = ccFixedSingle
    .ListViewTeeth.ListItems.Clear               '清空列表
    .ListViewTeeth.ColumnHeaders.Clear           '清空列表头
    .ListViewTeeth.View = lvwReport              '设置列表显示方式
    .ListViewTeeth.LabelEdit = lvwManual         '禁止标签编辑
    .ListViewTeeth.MultiSelect = False
    '.listviewteeth.FlatScrollBar = False         '显示滚动条
    SetExtendedStyle .ListViewTeeth, LVS_EX_GRIDLINES, 0 '显示网络线
    SetExtendedStyle .ListViewTeeth, LVS_EX_FULLROWSELECT, 0 '选择整行
    
    .ListViewTeeth.ColumnHeaders.Clear
    .ListViewTeeth.ColumnHeaders.Add , , "齿号", .ListViewTeeth.Width * 0.1
    .ListViewTeeth.ColumnHeaders.Add , , "齿类", .ListViewTeeth.Width * 0.15
    .ListViewTeeth.ColumnHeaders.Add , , "直径", .ListViewTeeth.Width * 0.15
    .ListViewTeeth.ColumnHeaders.Add , , "偏差", .ListViewTeeth.Width * 0.2
End With
End Sub

Sub InitListViewParameters()
With Form1
    .ListViewParameters.Appearance = ccFlat
    .ListViewParameters.BorderStyle = ccFixedSingle
    .ListViewParameters.ListItems.Clear               '清空列表
    .ListViewParameters.ColumnHeaders.Clear           '清空列表头
    .ListViewParameters.View = lvwReport              '设置列表显示方式
    .ListViewParameters.LabelEdit = lvwManual         '禁止标签编辑
    '.ListViewParameters.FlatScrollBar = False         '显示滚动条
    SetExtendedStyle .ListViewParameters, LVS_EX_GRIDLINES, 0 '显示网络线
    SetExtendedStyle .ListViewParameters, LVS_EX_FULLROWSELECT, 0 '选择整行
    
    .ListViewParameters.ColumnHeaders.Clear
    .ListViewParameters.ColumnHeaders.Add , , "序号", .ListViewParameters.Width * 0.05
    .ListViewParameters.ColumnHeaders.Add , , "参数名称", .ListViewParameters.Width * 0.48
    .ListViewParameters.ColumnHeaders.Add , , "变量", .ListViewParameters.Width * 0.13
    .ListViewParameters.ColumnHeaders.Add , , "值", .ListViewParameters.Width * 0.35
End With
End Sub
