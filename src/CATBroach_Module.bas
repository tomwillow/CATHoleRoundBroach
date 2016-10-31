Attribute VB_Name = "CATBroach_Module"
'CATBroach_Module.bas
Option Explicit
Option Base 1 '数组下界为1

'ShankVars------------------------------------------
Private D1#, Dq1#, D2#, D0#, D3#
Private L1#, lq1#, L2#, L3#, l0#, lq3#, l_l3#
Private C As Double

'CuttingVars----------------------------------------
Private Gammao#, Alphao#, Alphaoz#, balpha1_1#, balpha1_2#
        '前角γo，切削齿后角αo，校准齿后角αoz，切削齿刃带宽bα1_1，校准齿刃带宽bα1_2
Private Hasbalpha1 As Boolean
        '留刃带
Private po#, ho#, go#, l_ro#, U_Ro#, pz#, hz#, gz#, l_rz#, U_Rz#
Private nk%, bc#, hc#, rc#, Omegac#, DeltaAlphac#, HasChipDividingGroove As Boolean
Private DCutting() As Double
Private n1, n2, n3, n4, n As Integer
Private D4#, l_l4#, l_l#, lg#, lz#, L#

Public Sub oCreateRoundBroach()
SendMsgToParents "正在连接CATIA...请不要进行任何操作..."
If InitCATIAPart = True Then

    SendMsgToParents "正在生成柄部..."
    oInitShankVars
    oInitCuttingVars
    oCreateShank

    oCreateFrontPilot
    oCreateCutting
    oCreateRearPilot
    
    SendMsgToParents "完成！"
End If
End Sub

Private Sub SendMsgToParents(Msg As String)
Form2.Label1.Caption = Msg
Form1.SendMsgStr Msg
End Sub

Sub oInitShankVars()
D1 = Val(Form1.TextD1.Text)
Dq1 = Val(Form1.TextDq1.Text)
D2 = Val(Form1.TextD2.Text)
D0 = Val(Form1.TextD0.Text)
D3 = Val(Form1.TextD3.Text)

L1 = Val(Form1.TextL1.Text)
lq1 = 3
L2 = Val(Form1.TextL2.Text)
L3 = Val(Form1.TextU_L3.Text)
l0 = Val(Form1.Textl_l0.Text)
lq3 = Val(Form1.Combolq3.Text)
l_l3 = Val(Form1.Textl_l3.Text)

C = Val(Form1.TextC.Text)
End Sub

Sub oInitCuttingVars()
Dim i As Integer
Gammao = Val(Form1.TextGammaoDeg.Text) + Val(Form1.TextGammaoMin.Text) / 60 + Val(Form1.TextGammaoSec.Text) / 3600
Alphao = Val(Form1.TextAlphaoDeg.Text) + Val(Form1.TextAlphaoMin.Text) / 60 + Val(Form1.TextAlphaoSec.Text) / 3600
Alphaoz = Val(Form1.TextAlphaozDeg.Text) + Val(Form1.TextAlphaozMin.Text) / 60 + Val(Form1.TextAlphaozSec.Text) / 3600
If Form1.Checkbalpha1.Value = 1 Then
    Hasbalpha1 = True
    balpha1_1 = Val(Form1.Textbalpha1_1.Text)
    balpha1_2 = Val(Form1.Textbalpha1_2.Text)
Else
    Hasbalpha1 = False
    balpha1_1 = 0
    balpha1_2 = 0
End If

po = Val(Form1.Textp.Text)
ho = Val(Form1.Comboh.Text)
go = Val(Form1.Combog.Text)
l_ro = Val(Form1.Combol_r.Text)
U_Ro = Val(Form1.ComboU_R.Text)

pz = Val(Form1.Textpz.Text)
hz = Val(Form1.Combohz.Text)
gz = Val(Form1.Combogz.Text)
l_rz = Val(Form1.Combol_rz.Text)
U_Rz = Val(Form1.ComboU_Rz.Text)

ReDim DCutting(Form1.ListViewTeeth.ListItems.Count + 1) As Double
n1 = Val(Form1.Textn1.Text)
n2 = Val(Form1.Textn2.Text)
n3 = Val(Form1.Textn3.Text)
n4 = Val(Form1.Textn4.Text)
n = Val(Form1.ListViewTeeth.ListItems.Count)
For i = 1 To n
    DCutting(i) = Val(Form1.ListViewTeeth.ListItems(i).SubItems(2))
Next i
DCutting(n + 1) = DCutting(n)

l_l = po * (n1 + n2)
lg = pz * n3
lz = pz * n4
D4 = Val(Form1.TextD4.Text)
l_l4 = Val(Form1.Textl_l4.Text)

nk = Val(Form1.Textnk.Text)
bc = Val(Form1.Textbc.Text)
hc = Val(Form1.Texthc.Text)
rc = Val(Form1.Textrc.Text)
Omegac = Val(Form1.TextOmegac.Text)
DeltaAlphac = Val(Form1.TextDeltaAlphacDeg) + Val(Form1.TextDeltaAlphacMin) / 60 + Val(Form1.TextDeltaAlphacSec) / 3600
If Form1.CheckHasChipDividingGroove.Value = 1 Then
    HasChipDividingGroove = True
Else
    HasChipDividingGroove = False
End If
End Sub

Sub oCreateShank()
Dim i As Integer
Dim D0EquD1 As Boolean
D0EquD1 = Abs(D0 - D1) < 0.001

Dim oPlaneYZ As Plane '平面
Set oPlaneYZ = oPart.OriginElements.PlaneYZ

Dim oSketch As Sketch '草绘
Set oSketch = oBody.Sketches.Add(oPlaneYZ)

Dim oFactory2D As Factory2D
Set oFactory2D = oSketch.OpenEdition
    
    Dim oConstraints As Constraints '约束集
    Dim oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
    
    '设值-------------------------------------------------------------------------------
    Dim p(8, 2), p9a1x, p9a1y, p9a2x, p9a2y, p9a3x, p9a3y, p9b1x, p9b1y, p10x, p10y, p11x, p11y As Double
    p(1, 1) = 0: p(1, 2) = 0
    p(2, 1) = 0: p(2, 2) = D1 / 2
    p(3, 1) = L1: p(3, 2) = D1 / 2
    p(4, 1) = L1: p(4, 2) = Dq1 / 2
    p(5, 1) = L1 + lq1: p(5, 2) = Dq1 / 2
    p(6, 1) = L1 + lq1: p(6, 2) = D2 / 2
    p(7, 1) = L1 + L2: p(7, 2) = D2 / 2
    p(8, 1) = L1 + L2: p(8, 2) = D1 / 2
    p9a1x = L3: p9a1y = D1 / 2
    p9a2x = L3: p9a2y = D0 / 2
    p9a3x = L1 + L2 + l0 - lq3: p9a3y = D0 / 2
    p9b1x = L1 + L2 + l0 - lq3: p9b1y = D1 / 2
    p10x = L1 + L2 + l0: p10y = D3 / 2
    p11x = L1 + L2 + l0: p11y = 0
    
    '画点-------------------------------------------------------------------------------
    Dim op(8), op9a1, op9a2, op9a3, op9b1, op10, op11 As Point2D  '点1
    Set op(1) = oSketch.AbsoluteAxis.Origin
    For i = 2 To 8
        Set op(i) = oFactory2D.CreatePoint(p(i, 1), p(i, 2))
    Next i
    If Not D0EquD1 Then '颈部直径同D0不同
        Set op9a1 = oFactory2D.CreatePoint(p9a1x, p9a1y)
        Set op9a2 = oFactory2D.CreatePoint(p9a2x, p9a2y)
        Set op9a3 = oFactory2D.CreatePoint(p9a3x, p9a3y)
    Else
        Set op9b1 = oFactory2D.CreatePoint(p9b1x, p9b1y)
    End If
    Set op10 = oFactory2D.CreatePoint(p10x, p10y)
    Set op11 = oFactory2D.CreatePoint(p11x, p11y)
    
    '画线-------------------------------------------------------------------------------
    Dim oLine(7), oLine8a1, oLine8a2, oLine8a3, oLine8a4, oLine8b1, oLine8b2, oLine9 As Line2D
    'line1-7
    For i = 1 To 7
        Set oLine(i) = oFactory2D.CreateLine(p(i, 1), p(i, 2), p(i + 1, 1), p(i + 1, 2))
        oLine(i).StartPoint = op(i)
        oLine(i).EndPoint = op(i + 1)
    Next i
    'line8
    If Not D0EquD1 Then '颈部直径同D0不同
        Set oLine8a1 = oFactory2D.CreateLine(p(8, 1), p(8, 2), p9a1x, p9a1y)
        Set oLine8a2 = oFactory2D.CreateLine(p9a1x, p9a1y, p9a2x, p9a2y)
        Set oLine8a3 = oFactory2D.CreateLine(p9a2x, p9a2y, p9a3x, p9a3y)
        Set oLine8a4 = oFactory2D.CreateLine(p9a3x, p9a3y, p10x, p10y)
        oLine8a1.StartPoint = op(8)
        oLine8a1.EndPoint = op9a1
        oLine8a2.StartPoint = op9a1
        oLine8a2.EndPoint = op9a2
        oLine8a3.StartPoint = op9a2
        oLine8a3.EndPoint = op9a3
        oLine8a4.StartPoint = op9a3
        oLine8a4.EndPoint = op10
    Else
        Set oLine8b1 = oFactory2D.CreateLine(p(8, 1), p(8, 2), p9b1x, p9b1y)
        Set oLine8b2 = oFactory2D.CreateLine(p9b1x, p9b1y, p10x, p10y)
        oLine8b1.StartPoint = op(8)
        oLine8b1.EndPoint = op9b1
        oLine8b2.StartPoint = op9b1
        oLine8b2.EndPoint = op10
    End If
    'line9-10
    Set oLine9 = oFactory2D.CreateLine(p10x, p10y, p11x, p11y)
    oLine9.StartPoint = op10
    oLine9.EndPoint = op11
    
    Dim oCenterLine As Line2D
    Set oCenterLine = oFactory2D.CreateLine(p(1, 1), p(1, 2), p11x, p11y)
    oCenterLine.StartPoint = oSketch.AbsoluteAxis.Origin
    oCenterLine.EndPoint = op11
    oSketch.CenterLine = oCenterLine
    '水平，垂直约束-------------------------------------------------------------------------------
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(1)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(3)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(5)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(7)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine9
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine(2)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine(4)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine(6)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oCenterLine
    If Not D0EquD1 Then '颈部直径同D0不同
        oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine8a1
        oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine8a2
        oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine8a3
    Else
        oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine8b1
    End If
    
    '尺寸约束-------------------------------------------------------------------------------
    '                        标注直径
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine(2), oCenterLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine(4), oCenterLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine(6), oCenterLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, op10, oCenterLine
    
    '                        标注长度
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeLength, oLine(4)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeLength, oLine(6)
    
    '                        标注距离
    oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, oLine(1), oLine(5)
    oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, oLine(7), op10
    
    If Not D0EquD1 Then '颈部直径同D0不同
        oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine8a1, oCenterLine
        oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine8a3, oCenterLine
        oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, oLine(1), oLine8a2
        oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, op9a3, oLine9
    Else
        oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine8b1, oCenterLine
        oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, op9b1, oLine9
    End If
    
oSketch.CloseEdition

oPart.InWorkObject = oSketch

Dim oSF As ShapeFactory
Set oSF = oPart.ShapeFactory

'Dim RevoluteRef As Reference
'Set RevoluteRef = oPart.CreateReferenceFromObject(oSketch.AbsoluteAxis.HorizontalReference)

Dim oShaft As Shaft
Set oShaft = oSF.AddNewShaft(oSketch)
'oShaft.RevoluteAxis = RevoluteRef

oPart.InWorkObject = oShaft
oPart.Update

'倒角------------------------------------------------
Dim Str1, Str2, Str3 As String
If Not D0EquD1 Then '颈部直径同D0不同
    Str1 = "REdge:(Edge:(Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;13)));None:();Cf11:());Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;12)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR14)"
    Str2 = "REdge:(Edge:(Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;9)));None:();Cf11:());Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;8)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR14)"
    Str3 = "REdge:(Edge:(Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;8)));None:();Cf11:());Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;7)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR14)"
Else '默认相同
    Str1 = "REdge:(Edge:(Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;11)));None:();Cf11:());Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;10)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR14)"
    Str2 = "REdge:(Edge:(Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;7)));None:();Cf11:());Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;6)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR14)"
    Str3 = "REdge:(Edge:(Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;6)));None:();Cf11:());Face:(Brp:(Shaft.1;0:(Brp:(Sketch.1;5)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR14)"
End If
oCreateChamfer oShaft, Str1, C
oCreateChamfer oShaft, Str2, (Dq1 - D2) / 2
oCreateChamfer oShaft, Str3, (D1 - D2) / 2

End Sub

Sub oCreateCutting()
Dim sXin#, sDin#, sDout#, sp#, sh#, sg#, sU_R#, sl_r#, sGammao#, sAlphao#, sHasbalpha1 As Boolean, sbalpha1#, sisFinal As Boolean
Dim Lasti%, LastsXout#, LastsDout#, Lastsg#, LastsAlphao#, LastsHasChipDividingGroove As Boolean, Lastsbalpha1#
LastsHasChipDividingGroove = False
Dim sHasChipDividingGroove As Boolean
Dim i%, S As String
sHasChipDividingGroove = HasChipDividingGroove
sXin = L1 + L2 + l0 + l_l3 - (po - go)
For i = 1 To n
    '初始化不同参数------------------------------------------------------------------
    
    Select Case i
    Case Is = 1: '第1齿采用p,h,g,R,r,Gammao,Alphao,hasbalpha1,balpha1_1
        S = "粗切齿"
        sDin = D3: sDout = DCutting(i): sp = po: sh = ho: sg = go: sU_R = U_Ro: sl_r = l_ro: sGammao = Gammao: sAlphao = Alphao: sHasbalpha1 = Hasbalpha1: sbalpha1 = balpha1_1: sisFinal = False
    
    Case Is <= n1 + n2: '粗切齿，过渡齿：n1,n2采用p,h,g,R,r,Gammao,Alphao,balpha1_1
        If i <= n1 Then
            S = "粗切齿"
        Else
            S = "过渡齿"
        End If
        sDin = LastsDout - (Lastsg - Lastsbalpha1) * Tan(DegToRad(LastsAlphao)) * 2
        sDout = DCutting(i)
        sp = po: sh = ho: sg = go: sU_R = U_Ro: sl_r = l_ro: sGammao = Gammao: sAlphao = Alphao: sHasbalpha1 = Hasbalpha1: sbalpha1 = balpha1_1: sisFinal = False
    
    Case Is <= n1 + n2 + 1: '过渡齿-精切齿：p交界处,h,gz,R,r,Gammao,Alphao,balpha1_1
        S = "过渡齿"
        sDin = LastsDout - (Lastsg - Lastsbalpha1) * Tan(DegToRad(LastsAlphao)) * 2
        sDout = DCutting(i)
        sp = po - go + gz: sh = ho: sg = gz: sU_R = U_Ro: sl_r = l_ro: sGammao = Gammao: sAlphao = Alphao: sHasbalpha1 = Hasbalpha1: sbalpha1 = balpha1_1: sisFinal = False
    
    Case Is <= n1 + n2 + n3 '精切齿：n3采用pz,hz,gz,Rz,rz,Gammao,Alphao,balpha1_1
        S = "精切齿"
        sDin = LastsDout - (Lastsg - Lastsbalpha1) * Tan(DegToRad(LastsAlphao)) * 2
        sDout = DCutting(i)
        sp = pz: sh = hz: sg = gz: sU_R = U_Rz: sl_r = l_rz: sGammao = Gammao: sAlphao = Alphao: sHasbalpha1 = Hasbalpha1: sbalpha1 = balpha1_1: sisFinal = False
    
    Case Is <= n  '校准齿：n4-1采用pz,hz,gz,Rz,rz,Gammao,Alphaoz,balpha1_2
       S = "校准齿"
       sDin = LastsDout - (Lastsg - Lastsbalpha1) * Tan(DegToRad(LastsAlphao)) * 2
       sDout = DCutting(i)
       sp = pz: sh = hz: sg = gz: sU_R = U_Rz: sl_r = l_rz: sGammao = Gammao: sAlphao = Alphaoz: sHasbalpha1 = Hasbalpha1: sbalpha1 = balpha1_2: sisFinal = False
       sHasChipDividingGroove = False
            
    End Select
    
    SendMsgToParents "正在生成第" & i & "/" & n & "齿，类型：" & S & "..."
    
    '生成容屑槽
    oCreateOneCuttingFromU_R sXin, sDin, sDout, sp, sh, sg, sU_R, sl_r, sGammao, sAlphao, sHasbalpha1, sbalpha1, sisFinal
    
    '生成分屑槽：由于直接按i同步生成分屑槽时不能切除容屑槽第一曲面，故将分屑槽延后1步
    If LastsHasChipDividingGroove = True Then oCreateOneChipDividingGroove Lasti, LastsXout, LastsDout, Lastsg, LastsAlphao, DeltaAlphac, nk, bc, hc, rc, Omegac
    
    LastsHasChipDividingGroove = sHasChipDividingGroove
    Lasti = i
    LastsXout = sXin + sp
    LastsDout = sDout
    Lastsg = sg
    LastsAlphao = sAlphao
    '分屑槽延后参数计算完毕
    
    Lastsbalpha1 = sbalpha1
    
    '增量Xin计算
    sXin = sXin + sp
Next i
'最后1齿采用pz,hz,gz,Rz,rz,Gammao=0,Alphaoz,balpha1_2
SendMsgToParents "正在生成校准部-后导部过渡..."
sDin = LastsDout - (Lastsg - Lastsbalpha1) * Tan(DegToRad(LastsAlphao)) * 2
sDout = D4
sp = pz: sh = hz: sg = gz: sU_R = U_Rz: sl_r = l_rz: sGammao = 0: sAlphao = Alphaoz: sHasbalpha1 = Hasbalpha1: sbalpha1 = balpha1_2: sisFinal = True
oCreateOneCuttingFromU_R sXin, sDin, sDout, sp, sh, sg, sU_R, sl_r, sGammao, sAlphao, sHasbalpha1, sbalpha1, sisFinal

End Sub

Sub oCreateFrontPilot()
Dim i%
Dim oPlaneYZ As Plane '平面
Set oPlaneYZ = oPart.OriginElements.PlaneYZ

Dim oSketch As Sketch '草绘
Set oSketch = oBody.Sketches.Add(oPlaneYZ)

Dim oFactory2D As Factory2D
Set oFactory2D = oSketch.OpenEdition

    Dim oConstraints As Constraints '约束集
    Dim oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
    
    Dim opX(4) As Double, opY(4) As Double
    opX(1) = L1 + L2 + l0: opY(1) = 0
    opX(2) = opX(1): opY(2) = D3 / 2
    opX(3) = opX(1) + l_l3 - (po - go): opY(3) = opY(2)
    opX(4) = opX(3): opY(4) = 0
    
    Dim op(4) As Point2D
    Dim oLine(3) As Line2D
    
    '生成中心线---------------------------------------------------------------------
    Dim oCenterLine As Line2D
    Set oCenterLine = oFactory2D.CreateLine(opX(1), 0, opX(4), 0)
    oSketch.CenterLine = oCenterLine
    
    For i = 1 To 4
        Set op(i) = oFactory2D.CreatePoint(opX(i), opY(i))
    Next i
    
    For i = 1 To 3
        Set oLine(i) = oFactory2D.CreateLine(opX(i), opY(i), opX(i + 1), opY(i + 1))
        oLine(i).StartPoint = op(i)
        oLine(i).EndPoint = op(i + 1)
    Next i
    
    '水平-垂直
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oCenterLine
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine(2)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(1)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(3)
    
    '重合
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, oCenterLine, oSketch.AbsoluteAxis.Origin
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, op(1), oCenterLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, op(4), oCenterLine
    
    '距离
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine(2), oCenterLine
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeLength, oLine(2)
    oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, oLine(1), oSketch.AbsoluteAxis.VerticalReference
    
oSketch.CloseEdition

oPart.InWorkObject = oSketch

Dim oSF As ShapeFactory
Set oSF = oPart.ShapeFactory

Dim oShaft As Shaft
Set oShaft = oSF.AddNewShaft(oSketch)

oPart.InWorkObject = oShaft
oPart.Update
End Sub

Sub oCreateOneCuttingFromU_R(ByVal Xin#, ByVal Din#, ByVal Dout#, ByVal ips#, ByVal ih#, ByVal ig#, ByVal iU_R#, ByVal il_r#, _
                    ByVal iGammao#, ByVal iAlphao#, ByVal iHasbalpha1 As Boolean, ByVal ibalpha1#, ByVal isFinal As Boolean)
Dim i As Integer
Dim j As Integer
'初始化不同参数------------------------------------------------------------------
Dim iGammaoRad#, iAlphaoRad#, AuxLength#
iGammaoRad = DegToRad(iGammao)
iAlphaoRad = DegToRad(iAlphao)
AuxLength = ig

Dim oPlaneYZ As Plane '平面
Set oPlaneYZ = oPart.OriginElements.PlaneYZ

Dim oSketch As Sketch '草绘
Set oSketch = oBody.Sketches.Add(oPlaneYZ)

Dim oFactory2D As Factory2D
Set oFactory2D = oSketch.OpenEdition

    Dim oConstraints As Constraints '约束集
    Dim oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
    '中间段草图-----------------------------------------------------------------
    Dim opX(6) As Double, opY(6) As Double '点坐标
    Dim op(6) As Point2D '点
    
    Dim opCenterX(2) As Double, opCenterY(2) As Double '圆心坐标
    Dim opCenter(2) As Point2D '圆心点
    
    Dim opConX As Double, opConY As Double '辅助点坐标
    Dim opCon As Point2D '辅助点
    
    Dim oLine(3) As Line2D '线段
    
    Dim oCircle(2) As Circle2D '圆弧
    
    Dim oConLineH As Line2D
    Dim oConLineV As Line2D
    Dim oConLine As Line2D
    
    '生成中心线---------------------------------------------------------------------
    Dim oCenterLine As Line2D
    Set oCenterLine = oFactory2D.CreateLine(0, 0, Xin + ips, 0)
    oSketch.CenterLine = oCenterLine
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oCenterLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, oCenterLine, oSketch.AbsoluteAxis.Origin

    '计算点坐标---------------------------------------------------------------------
    opX(1) = Xin: opY(1) = Din / 2
    opX(4) = opX(1) + ips - ig: opY(4) = Dout / 2
    oCalcPointXY opX(1), opY(1), opX(4), opY(4), ips, ih, ig, iU_R, il_r, iGammaoRad, _
                 opX(2), opY(2), opX(3), opY(3), opCenterX(1), opCenterY(1), opCenterX(2), opCenterY(2), opConX, opConY
    opX(5) = opX(4) + ibalpha1: opY(5) = opY(4)
    opX(6) = opX(4) + ig: opY(6) = opY(4) - (ig - ibalpha1) * Tan(iAlphaoRad)
    '生成点---------------------------------------------------------------------
    For j = 1 To 6 '第1点非线段端点
        If (Hasbalpha1 = False) And (j = 5) Then
            GoTo EndofFor
        End If
        Set op(j) = oFactory2D.CreatePoint(opX(j), opY(j))
EndofFor:
    Next j
    Set opCenter(1) = oFactory2D.CreatePoint(opCenterX(1), opCenterY(1))
    Set opCenter(2) = oFactory2D.CreatePoint(opCenterX(2), opCenterY(2))
    Set opCon = oFactory2D.CreatePoint(opConX, opConY)
    
    '生成草图---------------------------------------------------------------------
    
    '弧线1
    Set oCircle(1) = oFactory2D.CreateCircle(opCenterX(1), opCenterY(1), iU_R, _
                                                GetAngleFromPoint(opCenterX(1), opCenterY(1), opX(1), opY(1)), _
                                                GetAngleFromPoint(opCenterX(1), opCenterY(1), opX(2), opY(2)))
    oCircle(1).StartPoint = op(1)
    oCircle(1).EndPoint = op(2)
    oCircle(1).CenterPoint = opCenter(1)
    
    '弧线2
    Set oCircle(2) = oFactory2D.CreateCircle(opCenterX(2), opCenterY(2), il_r, _
                                                GetAngleFromPoint(opCenterX(2), opCenterY(2), opX(2), opY(2)), _
                                                GetAngleFromPoint(opCenterX(2), opCenterY(2), opX(3), opY(3)))
    oCircle(2).StartPoint = op(2)
    oCircle(2).EndPoint = op(3)
    oCircle(2).CenterPoint = opCenter(2)
    
    '线段1
    Set oLine(1) = oFactory2D.CreateLine(opX(3), opY(3), opX(4), opY(4))
    oLine(1).StartPoint = op(3)
    oLine(1).EndPoint = op(4)
    If isFinal = True Then
        oLine(1).Construction = True
    End If
    
    '线段2-3
    If Not isFinal Then
        If iHasbalpha1 = True Then
            Set oLine(2) = oFactory2D.CreateLine(opX(4), opY(4), opX(5), opY(5))
            oLine(2).StartPoint = op(4)
            oLine(2).EndPoint = op(5)
            Set oLine(3) = oFactory2D.CreateLine(opX(5), opY(5), opX(6), opY(6))
            oLine(3).StartPoint = op(5)
            oLine(3).EndPoint = op(6)
        Else
            Set oLine(2) = oFactory2D.CreateLine(opX(4), opY(4), opX(6), opY(6))
            oLine(2).StartPoint = op(4)
            oLine(2).EndPoint = op(6)
        End If

        '辅助线V
        Set oConLineV = oFactory2D.CreateLine(opX(4), opY(4), opX(4), opY(4) + AuxLength)
        oConLineV.StartPoint = op(4)
        oConLineV.Construction = True

        '辅助线H
        Set oConLineH = oFactory2D.CreateLine(opX(4) - AuxLength, opY(4), opX(4), opY(4))
        oConLineH.EndPoint = op(4)
        oConLineH.Construction = True
    End If

    '辅助线
    Set oConLine = oFactory2D.CreateLine(opConX - AuxLength, opConY, opConX, opConY)
    oConLine.EndPoint = opCon
    oConLine.Construction = True
    '加入约束--------------------------------------------------------------------------

    '水平-垂直
    If Not isFinal Then
        If iHasbalpha1 = True Then
            oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine(2)
        End If
        oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oConLineV 'V辅助线垂直
        oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oConLineH 'H辅助线水平
    End If
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oConLine '辅助线水平

    '重合
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, opCon, oLine(1) '辅助线1端点位于直线1上

    '后角αo
    If Not isFinal Then
        If iHasbalpha1 = True Then
            oAddBiEltCst oConstraints, oConstraint, catCstTypeAngle, oLine(2), oLine(3)
            oConstraint.Dimension.Value = iAlphao
        Else
            oAddBiEltCst oConstraints, oConstraint, catCstTypeAngle, oLine(2), oConLineH
            oConstraint.Dimension.Value = iAlphao
        End If
    End If

    '前角γ
    If Not isFinal Then
        oAddBiEltCst oConstraints, oConstraint, catCstTypeAngle, oLine(1), oConLineV
        oConstraint.Dimension.Value = iGammao
    End If

    '圆弧半径R,r
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeRadius, oCircle(1)
    oConstraint.Dimension.Value = iU_R
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeRadius, oCircle(2)
    oConstraint.Dimension.Value = il_r

    '相切
    oAddBiEltCst oConstraints, oConstraint, catCstTypeTangency, oCircle(1), oCircle(2)  '弧1与弧2相切
    oAddBiEltCst oConstraints, oConstraint, catCstTypeTangency, oCircle(2), oLine(1) '弧2与直线1相切
    oAddBiEltCst oConstraints, oConstraint, catCstTypeTangency, oCircle(2), oConLine '弧2与辅助线1相切

    '距离
    oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, op(1), oSketch.AbsoluteAxis.VerticalReference   'Xin
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, op(1), oCenterLine 'Din
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, op(4), oCenterLine 'Dout

    If Not isFinal Then 'p,g
        oAddTriEltCst oConstraints, oConstraint, catCstTypeDistance, op(4), op(6), oSketch.AbsoluteAxis.HorizontalReference
        oConstraint.Dimension.Value = ig
        oAddTriEltCst oConstraints, oConstraint, catCstTypeDistance, op(1), op(6), oSketch.AbsoluteAxis.HorizontalReference
        oConstraint.Dimension.Value = ips
    Else
        oAddTriEltCst oConstraints, oConstraint, catCstTypeDistance, op(1), op(3), oSketch.AbsoluteAxis.HorizontalReference
        oConstraint.Dimension.Value = ips - ig
    End If
    oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, op(4), oConLine 'h
    oConstraint.Dimension.Value = ih

    If Not isFinal Then
        If iHasbalpha1 = True Then
            oAddMonoEltCst oConstraints, oConstraint, catCstTypeLength, oLine(2) 'bα1
            oConstraint.Dimension.Value = ibalpha1
        End If
    End If

    '中间段草图生成完毕---------------------------------------------------------

    '生成两端封闭线---------------------------------------------------------
    Dim opAX#, opAY#, opBX#, opBY#
    opAX = opX(1): opAY = 0
    If Not isFinal Then
        opBX = opX(6): opBY = 0
    Else
        opBX = opX(3): opBY = 0
    End If

    Dim opA As Point2D, opB As Point2D
    Set opA = oFactory2D.CreatePoint(opAX, opAY)
    Set opB = oFactory2D.CreatePoint(opBX, opBY)

    Dim oALine As Line2D
    Set oALine = oFactory2D.CreateLine(opAX, opAY, opX(1), opY(1))
    oALine.StartPoint = opA
    oALine.EndPoint = op(1)
    
    Dim oBLine As Line2D
    If Not isFinal Then
        Set oBLine = oFactory2D.CreateLine(opBX, opBY, opX(6), opY(6))
        oBLine.StartPoint = opB
        oBLine.EndPoint = op(6)
    Else
        Set oBLine = oFactory2D.CreateLine(opBX, opBY, opX(3), opY(3))
        oBLine.StartPoint = opB
        oBLine.EndPoint = op(3)
    End If

    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oALine
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oBLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, opA, oCenterLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, opB, oCenterLine
    If isFinal Then
        oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(1)
    End If
    
oSketch.CloseEdition

oPart.InWorkObject = oSketch

Dim oSF As ShapeFactory
Set oSF = oPart.ShapeFactory

Dim oShaft As Shaft
Set oShaft = oSF.AddNewShaft(oSketch)

oPart.InWorkObject = oShaft
oPart.Update
End Sub
Sub oCreateOneChipDividingGroove(ByVal iIndex%, ByVal iXout#, ByVal iDout#, ByVal ig#, ByVal iAlpha#, ByVal iDeltaAlphac#, _
                                   ByVal ink%, ByVal ibc#, ByVal ihc#, ByVal irc#, ByVal iOmegac#)
'生成斜面法线------------------------------------------------------
Dim oPlane As Plane
Set oPlane = oPart.OriginElements.PlaneYZ

Dim oSketch As Sketch '草绘
Set oSketch = oBody.Sketches.Add(oPlane)
'oSketch.Name = "Sketch.4"

Dim oFactory2D As Factory2D
Set oFactory2D = oSketch.OpenEdition

    Dim oConstraints As Constraints '约束集
    Dim oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
    
    Dim opX1#, opY1#, opX2#, opY2#
    opX1 = iXout - ig: opY1 = iDout / 2 + ig * Tan(DegToRad(iAlpha + iDeltaAlphac))
    opX2 = iXout: opY2 = iDout / 2
    
    Dim op1 As Point2D, op2 As Point2D
    Set op1 = oFactory2D.CreatePoint(opX1, opY1)
    Set op2 = oFactory2D.CreatePoint(opX2, opY2)
    
    Dim oCenterLine As Line2D
    Set oCenterLine = oFactory2D.CreateLine(opX1, 0, opX2, 0)
    oSketch.CenterLine = oCenterLine
    
    Dim oLine As Line2D
    Set oLine = oFactory2D.CreateLine(opX1, opY1, opX2, opY2)
    oLine.StartPoint = op1
    oLine.EndPoint = op2
    
    op1.Name = "Point1"
    'On
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, oCenterLine, oSketch.AbsoluteAxis.Origin
    
    '垂直-水平
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oCenterLine
    
    '角度
    oAddBiEltCst oConstraints, oConstraint, catCstTypeAngle, oLine, oCenterLine
    
    '距离
    oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, op2, oSketch.AbsoluteAxis.VerticalReference
    oAddTriEltCst oConstraints, oConstraint, catCstTypeDistance, op1, op2, oSketch.AbsoluteAxis.HorizontalReference
    
    '直径
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, op2, oCenterLine
    
oSketch.CloseEdition

oPart.InWorkObject = oSketch

'生成斜面---------------------------------------------------------------------------------------
Dim ohybridShapeFactory As HybridShapeFactory
Set ohybridShapeFactory = oPart.HybridShapeFactory

Dim oHybridShapePointCoord As HybridShapePointCoord
Set oHybridShapePointCoord = oPart.HybridShapeFactory.AddNewPointCoord(0, opX1, opY1)

Dim ref1 As Reference, ref2 As Reference
Set ref1 = oPart.CreateReferenceFromObject(oSketch)
Set ref2 = oPart.CreateReferenceFromObject(oHybridShapePointCoord)
'Set ref2 = oPart.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(" & oSketch.Name & ";2);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", oSketch)

Dim oHybridShapePlaneNormal As HybridShapePlaneNormal
Set oHybridShapePlaneNormal = ohybridShapeFactory.AddNewPlaneNormal(ref1, ref2)

oBody.InsertHybridShape oHybridShapePlaneNormal

oPart.InWorkObject = oHybridShapePlaneNormal
oPart.Update

'在生成平面上拉伸切除-----------------------------------------------------------------------
Dim oHybridBodies As HybridBodies
Set oHybridBodies = oBody.HybridBodies

'Dim oPlaneRef As Reference
'Set oPlaneRef = oPart.CreateReferenceFromObject(oHybridShapePlaneNormal)

Dim oSketchChip As Sketch '草绘
Set oSketchChip = oBody.Sketches.Add(oHybridShapePlaneNormal)

Dim oFactory2DChip As Factory2D
Set oFactory2DChip = oSketchChip.OpenEdition()

    Dim oConstraints2 As Constraints '约束集
    Dim oConstraint2 As Constraint
    Set oConstraints2 = oSketchChip.Constraints
    
    Dim op1LX#, op1LY#, op2LX#, op2LY#, op1RX#, op1RY#, op2RX#, op2RY#, opCX#, opCY#
    op1LX = ibc / 2: op1LY = 0 '正确
    op1RX = -ibc / 2: op1RY = 0 '正确
    
    op2LX = irc * Cos(DegToRad(iOmegac / 2)): op2LY = ihc + irc * (Sin(DegToRad(iOmegac / 2)) - 1)
    op2RX = -op2LX: op2RY = op2LY '正确
    
    opCX = 0: opCY = ihc - irc '正确
    
    Dim op1L As Point2D, op1R As Point2D, op2L As Point2D, op2R As Point2D, opC As Point2D
    Set op1L = oFactory2DChip.CreatePoint(op1LX, op1LY)
    Set op1R = oFactory2DChip.CreatePoint(op1RX, op1RY)
    Set op2L = oFactory2DChip.CreatePoint(op2LX, op2LY)
    Set op2R = oFactory2DChip.CreatePoint(op2RX, op2RY)
    Set opC = oFactory2DChip.CreatePoint(opCX, opCY)
    
    Dim oCenterLine2 As Line2D
    Set oCenterLine2 = oFactory2DChip.CreateLine(0, 0, 0, ihc)
    oSketchChip.CenterLine = oCenterLine2
    
    Dim oLine1 As Line2D, oLine2L As Line2D, oLine2R As Line2D, oCircle1 As Circle2D
    Set oLine1 = oFactory2DChip.CreateLine(op1LX, op1LY, op1RX, op1RY)
    oLine1.StartPoint = op1L
    oLine1.EndPoint = op1R
    Set oLine2L = oFactory2DChip.CreateLine(op1LX, op1LY, op2LX, op2LY)
    oLine2L.StartPoint = op1L
    oLine2L.EndPoint = op2L
    Set oLine2R = oFactory2DChip.CreateLine(op1RX, op1RY, op2RX, op2RY)
    oLine2R.StartPoint = op1R
    oLine2R.EndPoint = op2R
    Set oCircle1 = oFactory2DChip.CreateCircle(opCX, opCY, irc, GetAngleFromPoint(opCX, opCY, op2LX, op2LY), GetAngleFromPoint(opCX, opCY, op2RX, op2RY))
    oCircle1.CenterPoint = opC
    oCircle1.StartPoint = op2L
    oCircle1.EndPoint = op2R
    
oSketchChip.CloseEdition

oPart.InWorkObject = oSketchChip

Dim oSF As ShapeFactory
Set oSF = oPart.ShapeFactory

Dim oPocket As Pocket
Set oPocket = oSF.AddNewPocket(oSketchChip, 20#)

oPocket.DirectionOrientation = catRegularOrientation
oPocket.FirstLimit.LimitMode = catUpToNextLimit

oPart.InWorkObject = oPocket
oPart.Update
'拉伸切除完毕---------------------------------------------------------------

'生成阵列-------------------------------------------------------------------
Dim oRefAxis As Reference
Set oRefAxis = oPart.CreateReferenceFromObject(oPart.OriginElements.PlaneZX)

Dim oCircPattern As CircPattern
Set oCircPattern = oPart.ShapeFactory.AddNewCircPattern(oPocket, 1, ink, 0, 360 / ink, 1, 1, Nothing, oRefAxis, True, 0#, True)

oPart.InWorkObject = oCircPattern
oPart.Update
'阵列完毕---------------------------------------------------------------

'生成旋转-------------------------------------------------------------------
Dim DegRotate#
If iIndex Mod 2 = 0 Then
    DegRotate = 180 / ink
Else
    DegRotate = -180 / ink
End If
Dim oRotate As Rotate
Set oRotate = oPart.ShapeFactory.AddNewRotate2(oRefAxis, DegRotate)

oPart.InWorkObject = oRotate
oPart.Update
End Sub

Sub oCreateRearPilot()
Dim i%
Dim oPlaneYZ As Plane '平面
Set oPlaneYZ = oPart.OriginElements.PlaneYZ

Dim oSketch As Sketch '草绘
Set oSketch = oBody.Sketches.Add(oPlaneYZ)

Dim oFactory2D As Factory2D
Set oFactory2D = oSketch.OpenEdition

    Dim oConstraints As Constraints '约束集
    Dim oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
    
    Dim opX(4) As Double, opY(4) As Double
    opX(1) = L1 + L2 + l0 + l_l3 + l_l + lg + lz: opY(1) = 0
    opX(2) = opX(1): opY(2) = D4 / 2
    opX(3) = opX(1) + l_l4: opY(3) = opY(2)
    opX(4) = opX(3): opY(4) = 0
    
    Dim op(4) As Point2D
    Dim oLine(3) As Line2D
    
    '生成中心线---------------------------------------------------------------------
    Dim oCenterLine As Line2D
    Set oCenterLine = oFactory2D.CreateLine(opX(1), 0, opX(4), 0)
    oSketch.CenterLine = oCenterLine
    
    For i = 1 To 4
        Set op(i) = oFactory2D.CreatePoint(opX(i), opY(i))
    Next i
    
    For i = 1 To 3
        Set oLine(i) = oFactory2D.CreateLine(opX(i), opY(i), opX(i + 1), opY(i + 1))
        oLine(i).StartPoint = op(i)
        oLine(i).EndPoint = op(i + 1)
    Next i
    
    '水平-垂直
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oCenterLine
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeHorizontality, oLine(2)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(1)
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeVerticality, oLine(3)
    
    '重合
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, oCenterLine, oSketch.AbsoluteAxis.Origin
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, op(1), oCenterLine
    oAddBiEltCst oConstraints, oConstraint, catCstTypeOn, op(4), oCenterLine
    
    '距离
    oAddBiEltCst oConstraints, oConstraint, catCstTypeCylinderRadius, oLine(2), oCenterLine
    oAddMonoEltCst oConstraints, oConstraint, catCstTypeLength, oLine(2)
    oAddBiEltCst oConstraints, oConstraint, catCstTypeDistance, oLine(1), oSketch.AbsoluteAxis.VerticalReference
    
oSketch.CloseEdition

oPart.InWorkObject = oSketch

Dim oSF As ShapeFactory
Set oSF = oPart.ShapeFactory

Dim oShaft As Shaft
Set oShaft = oSF.AddNewShaft(oSketch)

oPart.InWorkObject = oShaft
oPart.Update

Dim sShaft As String, sSketch As String, S As String
sShaft = FixBRepName(oShaft.Name)
sSketch = FixBRepName(oSketch.Name)

S = "REdge:(Edge:(Face:(Brp:(" & _
                        sShaft & _
                        ";0:(Brp:(" & _
                        sSketch & _
                        ";3)));None:();Cf11:());Face:(Brp:(" & _
                        sShaft & _
                        ";0:(Brp:(" & _
                        sSketch & _
                        ";1)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"

oCreateChamfer oShaft, S, C

End Sub

'整型可以用“%”代替，长整型可以用“&”代替，实型可以用“!”，双精度实型可以用“#”定义
Sub oCalcPointXY(ByVal opXi3#, ByVal opYi3#, ByVal opXi6#, ByVal opYi6#, ByVal p#, ByVal h#, ByVal g#, ByVal U_R#, ByVal l_r#, ByVal GammaoRad#, _
                 ByRef opXi4, opYi4, opXi5, opYi5, opCenterXi1, opCenterYi1, opCenterXi2, opCenterYi2, opConXi1, opConYi1)
Dim IJ, AH, AO2, theta1, theta2, theta3 As Double

opConXi1 = opXi6 + h * Tan(GammaoRad): opConYi1 = opYi6 - h
opCenterXi2 = opConXi1 - l_r / Tan((PI - 2 * GammaoRad) / 4): opCenterYi2 = opYi6 - h + l_r
opXi5 = opConXi1 - Sin(GammaoRad) * (l_r / Tan((PI - 2 * GammaoRad) / 4)): opYi5 = opConYi1 + Cos(GammaoRad) * (l_r / Tan((PI - 2 * GammaoRad) / 4))
IJ = p - g + h * Tan(GammaoRad) - l_r / Tan(PI / 4 - GammaoRad / 2)
AH = h - (opYi6 - opYi3) - l_r
AO2 = Sqr(AH ^ 2 + IJ ^ 2)
theta1 = ACos((AO2 ^ 2 + (U_R - l_r) ^ 2 - U_R ^ 2) / (2 * AO2 * (U_R - l_r)))
theta3 = Atn(AH / IJ)
theta2 = PI - theta3 - theta1
opXi4 = opCenterXi2 - l_r * Cos(theta2)
opYi4 = opCenterYi2 - l_r * Sin(theta2)
opCenterXi1 = opCenterXi2 + (U_R - l_r) * Cos(theta2)
opCenterYi1 = opCenterYi2 + (U_R - l_r) * Sin(theta2)
End Sub

