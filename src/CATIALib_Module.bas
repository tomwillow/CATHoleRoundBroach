Attribute VB_Name = "CATIALib_Module"
Option Explicit
' ***********************************************************************
'   目的：       标准功能模块
'   原作者：     SUNNYTECH Huting <tianshuen@gmail.com>
'   改动：       TomWillow
'   编程语言:    VB
'   语言:        中文
'   CATIA Level: V5R9
' ***********************************************************************

' --------------------------------------------------------------
' 窗口属性设定API声明
' --------------------------------------------------------------
Private Declare Function SetWindowPos Lib "User32" ( _
                                ByVal hWnd As Long, _
                                ByVal hWndInsertAfter As Long, _
                                ByVal X As Long, ByVal Y As Long, _
                                ByVal cx As Long, ByVal cy As Long, _
                                ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

' --------------------------------------------------------------
' 基本变量定义
' --------------------------------------------------------------
Public CATIA As INFITF.Application
Public oProductDoc As ProductDocument
Public oPartDoc As PartDocument
Public oDrawingDoc As DrawingDocument
Public oPart As Part

Public oBodies As Bodies
Public oBody As Body
Public oHBodies As HybridBodies
Public oHBody As HybridBody

Public oSF As ShapeFactory
Public oHSF As HybridShapeFactory

' ***********************************************************************
'   目的：      初始化CATIA产品文档，并初始化必要的基本变量
'
'   输入：      bNewProduct:   初始化时是否新建产品文件
'                              可选，默认新建文件
'               strProduct:    初始化时是否打开已经存在的产品文件
'                              可选，默认新建文件
' ***********************************************************************
Sub InitCATIAProduct(Optional bNewProduct As Boolean = True, _
                     Optional strProduct As String = "")
    
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
      Set CATIA = CreateObject("CATIA.Application")
      CATIA.Visible = True
    End If
    
    If bNewProduct Then
        Set oProductDoc = CATIA.Documents.Add("Product")
    Else
        If strProduct = "" Then
            Set oProductDoc = CATIA.ActiveDocument
            If oProductDoc Is Nothing Then
                Err.Clear
                Set oProductDoc = CATIA.Documents.Add("Product")
            End If
        Else
            If Dir(strProduct) <> "" Then
                Set oProductDoc = CATIA.Documents.Open(strProduct)
            Else
                MsgBox "指定的文件不存在！"
                End
            End If
        End If
    End If
    
    On Error GoTo 0

End Sub

' ***********************************************************************
'   目的：      初始化CATIA零件文档，并初始化必要的基本变量
'
'   输入：       bNewPart:    初始化时是否新建零件文件
'                             可选，默认新建文件
'                strPart:     初始化时是否打开已经存在的零件文件
'                             可选，默认新建文件
' ***********************************************************************
Function InitCATIAPart(Optional bNewPart As Boolean = True, _
                  Optional strPart As String = "") As Boolean

    'On Error GoTo out
    On Error Resume Next
    
    Set CATIA = GetObject(, "CATIA.Application") '连接到CATIA
    
    If Err.Number <> 0 Then 'CATIA未打开→打开CATIA
      Set CATIA = CreateObject("CATIA.Application")
      CATIA.Visible = True
    End If
    
    If CATIA.Caption <> "" Then '正常打开CATIA
    
        If bNewPart Then
            Set oPartDoc = CATIA.Documents.Add("Part")
        Else
            If strPart = "" Then
                Set oPartDoc = CATIA.ActiveDocument
                If oPartDoc Is Nothing Then
                    Err.Clear
                    Set oPartDoc = CATIA.Documents.Add("Part")
                End If
            Else
                If Dir(strPart) <> "" Then
                    Set oPartDoc = CATIA.Documents.Open(strPart)
                Else
                    MsgBox "指定的文件不存在！"
                    End
                End If
            End If
        End If
        
        Set oPart = oPartDoc.Part
        Set oBodies = oPart.Bodies
        Set oBody = oPart.MainBody
        Set oHBodies = oPart.HybridBodies
        
        Set oSF = oPart.ShapeFactory
        Set oHSF = oPart.HybridShapeFactor
        
        InitCATIAPart = True '返回正常信号
        On Error GoTo 0
    Else
        MsgBox "未检测到CATIA。"
        Form1.SendMsgStr "未检测到CATIA。"
    End If
        

End Function

' ***********************************************************************
'   目的：      初始化CATIA工程图文档，并初始化必要的基本变量
'
'   输入：       bNewDrawing:    初始化时是否新建零件文件
'                                可选，默认新建文件
'                strDrawing:     初始化时是否打开已经存在的工程图文件
'                                可选，默认新建文件
' ***********************************************************************
Sub InitCATIADrawing(Optional bNewDrawing As Boolean = True, _
                     Optional strDrawing As String = "")

    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
      Set CATIA = CreateObject("CATIA.Application")
      CATIA.Visible = True
    End If
    
    If bNewDrawing Then
        Set oDrawingDoc = CATIA.Documents.Add("Drawing")
    Else
        If bNewDrawing = "" Then
            Set oDrawingDoc = CATIA.ActiveDocument
            If oDrawingDoc Is Nothing Then
                Err.Clear
                Set oDrawingDoc = CATIA.Documents.Add("Drawing")
            End If
        Else
            If Dir(bNewDrawing) <> "" Then
                Set oDrawingDoc = CATIA.Documents.Open(strDrawing)
            Else
                MsgBox "指定的文件不存在！"
                End
            End If
        End If
    End If
    
    On Error GoTo 0

End Sub

' ***********************************************************************
'   目的：      设置窗口使其始终在其它窗口上面
'
'   输入：      iHwnd:    要设置的窗口句柄
'               bOnTop:   设置或取消窗口的置顶属性
'                         可选，默认为真
' ***********************************************************************
Sub MakeMeOnTop(iHwnd As Long, Optional bOnTop As Boolean = True)
    
    If bOnTop Then
        SetWindowPos iHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Else
        SetWindowPos iHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    End If

End Sub

' ***********************************************************************
'   目的：      添加新的几何元素集
'
'   输入：      HBodyName: 几何元素集名称
' ***********************************************************************
Function AddHBody(Optional HBodyName As String = "") As HybridBody
    
    Dim oHB As HybridBody
    
    On Error Resume Next
    
    Set oHB = oHBodies.Add()
    If HBodyName <> "" Then
        oHB.Name = HBodyName
    End If
    
    Set AddHBody = oHB
    
    On Error GoTo 0
    
End Function

' ***********************************************************************
'   目的：      隐藏/显示元素
'
'   输入：      Element: 要隐藏/显示的元素
'               isShow:  要隐藏或显示该元素
'                        可选，默认隐藏
' ***********************************************************************
Sub HideShow(Element, Optional isShow As Boolean = False)
    
    Dim RefElement As Reference
    
    Set RefElement = oPart.CreateReferenceFromObject(Element)
    oHSF.GSMVisibility RefElement, isShow
    
End Sub

Public Function FixBRepName(ByVal iBRepName As String) As String
FixBRepName = iBRepName
FixBRepName = Replace(FixBRepName, "旋转体", "Shaft")
FixBRepName = Replace(FixBRepName, "草图", "Sketch")
End Function

' ***********************************************************************
'   目的：      生成倒角
'
'   输入：      Object1: 引用传递。需倒角几何体
'               iLabel: 元素标记
'               oLength:倒角长度
' ***********************************************************************
Sub oCreateChamfer(ByRef Object1 As AnyObject, ByVal iLabel As String, oLength As Double)
Dim ref1, ref2 As Reference
Set ref1 = oPart.CreateReferenceFromName("")
Set ref2 = oPart.CreateReferenceFromBRepName(iLabel, Object1)

Dim oSF As ShapeFactory
Set oSF = oPart.ShapeFactory

Dim oChamfer1 As Chamfer
Set oChamfer1 = oSF.AddNewChamfer(ref1, catTangencyChamfer, catLengthAngleChamfer, catNoReverseChamfer, 1#, 45#)

Dim olength1 As Length
Set olength1 = oChamfer1.Length1
olength1.Value = oLength

oChamfer1.AddElementToChamfer ref2
oChamfer1.Mode = catLengthAngleChamfer
oChamfer1.Propagation = catTangencyChamfer
oChamfer1.Orientation = catNoReverseChamfer

oPart.Update
End Sub

Function GetRef(iObject As AnyObject) As Reference
GetRef = oPart.CreateReferenceFromObject(iObject)
End Function
' ***********************************************************************
'   目的：      为一个元素添加约束
'
'   输入：      oConstraints: 约束集
'               iCstType: 约束类型
'               Object1,Object2: 元素
' ***********************************************************************
Sub oAddMonoEltCst(ByRef oConstraints As Constraints, ByRef oConstraint As Constraint, ByVal iCstType As CatConstraintType, ByVal Object1 As AnyObject, Optional ByVal Num As Double)

Dim oref1 As Reference
Set oref1 = oPart.CreateReferenceFromObject(Object1)

'Dim oconstraint As Constraint
Set oConstraint = oConstraints.AddMonoEltCst(iCstType, oref1)
'Dim ilength As Dimension
'Set ilength = oconstraint1.Dimension
'ilength.value = Num
'oconstraint1.Dimension.value = Num

End Sub

' ***********************************************************************
'   目的：      为两个元素添加约束
'
'   输入：      oConstraints: 约束集
'               iCstType: 约束类型
'               Object1,Object2: 元素
' ***********************************************************************
Sub oAddBiEltCst(ByRef oConstraints As Constraints, ByRef oConstraint As Constraint, ByVal iCstType As CatConstraintType, ByVal Object1 As AnyObject, ByVal Object2 As AnyObject, Optional ByVal Num As Double)

Dim oref1, oref2 As Reference
Set oref1 = oPart.CreateReferenceFromObject(Object1)
Set oref2 = oPart.CreateReferenceFromObject(Object2)

Set oConstraint = oConstraints.AddBiEltCst(iCstType, oref1, oref2)
'oConstraint1.Dimension.value = Num

End Sub

' ***********************************************************************
'   目的：      为三个元素添加约束
'
'   输入：      oConstraints: 约束集
'               iCstType: 约束类型
'               Object1,Object2: 元素
' ***********************************************************************
Sub oAddTriEltCst(ByRef oConstraints As Constraints, ByRef oConstraint As Constraint, ByVal iCstType As CatConstraintType, ByVal Object1 As AnyObject, ByVal Object2 As AnyObject, ByVal object3 As AnyObject, Optional ByVal Num As Double)

Dim oref1, oref2, oref3 As Reference
Set oref1 = oPart.CreateReferenceFromObject(Object1)
Set oref2 = oPart.CreateReferenceFromObject(Object2)
Set oref3 = oPart.CreateReferenceFromObject(object3)

Set oConstraint = oConstraints.AddTriEltCst(iCstType, oref1, oref2, oref3)
'oConstraint1.Dimension.value = Num

End Sub

