Option Explicit

Dim Connection1 As New ADODB.Connection
Dim Recordset1 As New ADODB.Recordset
Dim DatabaseName As String
Dim pX As Single, pY As Single
Dim OnEditToothIndex As Integer '正在编辑的齿号

Public Sub SendMsgStr(Msg As String)
TextMemo.Text = Now & ":" & Msg & Chr(13) & Chr(10) & TextMemo.Text
'StatusBar1.Panels(1).Text = Now & ":" & Msg
End Sub

Public Sub SetComboFromStr(ByRef Combo As ComboBox, ByVal S As String) '由ComboBox和列名称调整ComboBox的ListIndex
Dim i%
For i = 0 To Combo.ListCount - 1
    If Combo.List(i) = S Then Combo.ListIndex = i
Next i
End Sub

Private Sub SetOptionEnable() '根据拉削余量查表结果自动禁用没有的加工方式选项
Dim D As Double
Dim i As Integer
Dim StrPreHole As String
Recordset1.Open "SELECT * FROM 6_8圆孔拉削余量 WHERE D1<=" & D & " AND " & D & "<=D2 ORDER BY D1"
'选择适合的直径

OptionReaming.Enabled = False
OptionCounterboring.Enabled = False
OptionDrilling.Enabled = False
OptionBoring.Enabled = False
For i = 1 To Recordset1.RecordCount
'MsgBox AscW(Recordset1.Fields("预制孔加工方式")) & AscW("铰")
  If AscW(Recordset1.Fields("预制孔加工方式")) = AscW("铰") Then
    OptionReaming.Enabled = True
    OptionReaming.Value = True
    StrPreHole = "铰"
  End If
  If AscW(Recordset1.Fields("预制孔加工方式")) = AscW("扩") Then
    OptionCounterboring.Enabled = True
    OptionCounterboring.Value = True
    StrPreHole = "扩"
  End If
  If AscW(Recordset1.Fields("预制孔加工方式")) = AscW("钻") Then
    OptionDrilling.Enabled = True
    OptionDrilling.Value = True
    StrPreHole = "钻"
  End If
  If AscW(Recordset1.Fields("预制孔加工方式")) = AscW("镗") Then
    OptionBoring.Enabled = True
    OptionBoring.Value = True
    StrPreHole = "镗"
  End If
  Recordset1.MoveNext
Next i
Recordset1.Close
End Sub

Public Function GetStrFromMachiningMode() As String
If OptionReaming.Value = True Then GetStrFromMachiningMode = OptionReaming.Caption
If OptionCounterboring.Value = True Then GetStrFromMachiningMode = OptionCounterboring.Caption
If OptionDrilling.Value = True Then GetStrFromMachiningMode = OptionDrilling.Caption
If OptionBoring.Value = True Then GetStrFromMachiningMode = OptionBoring.Caption
End Function

Public Sub SetMachiningModeFromStr(ByVal S As String)
If S = OptionReaming.Caption Then OptionReaming.Value = True
If S = OptionCounterboring.Caption Then OptionCounterboring.Value = True
If S = OptionDrilling.Caption Then OptionDrilling.Value = True
If S = OptionBoring.Caption Then OptionBoring.Value = True
End Sub

Public Function GetStrFromSmallerOption() As String
If OptionSmallerGroove.Value = True Then GetStrFromSmallerOption = OptionSmallerGroove.Caption
If OptionNoSmallerGroove.Value = True Then GetStrFromSmallerOption = OptionNoSmallerGroove.Caption
End Function

Public Sub SetSmallerOptionFromStr(ByVal S As String)
If S = OptionSmallerGroove.Caption Then OptionSmallerGroove.Value = True
If S = OptionNoSmallerGroove.Caption Then OptionNoSmallerGroove.Value = True
End Sub

Public Function GetStrFromGrooveOption() As String
Dim i%
For i = 0 To 3
    If OptionGroove(i).Value = True Then GetStrFromGrooveOption = OptionGroove(i).Caption
Next i
End Function

Public Sub SetGrooveOptionFromStr(ByVal S As String)
Dim i%
For i = 0 To 3
    If S = OptionGroove(i).Caption Then OptionGroove(i).Value = True
Next i
End Sub

'拉削余量（公式）----------------------------------------------------------------------
Private Function GetA0FromFormula() As Double
If OptionReaming.Value = True Or OptionBoring.Value = True Then '铰或镗
    GetA0FromFormula = 0.005 * ValF(TextD.Text, "拉削直径D") + 0.05 * Sqr(ValF(TextL0.Text, "拉削长度L0"))
    Else
    If OptionCounterboring.Value = True Then '扩
        GetA0FromFormula = 0.005 * ValF(TextD.Text, "拉削直径D") + 0.075 * Sqr(ValF(TextL0.Text, "拉削长度L0"))
        Else
        If OptionDrilling.Value = True Then '钻
        GetA0FromFormula = 0.005 * ValF(TextD.Text, "拉削直径D") + 0.1 * Sqr(ValF(TextL0.Text, "拉削长度L0"))
        End If
    End If
End If
End Function

'拉削余量（查表）---------------------------------------------------------------
Private Function GetA0FromTable() As String
Dim i, n As Integer
Dim StrPreHole As String '预制孔加工方式
Dim ly(100), lydata(100), lymin, lymax As Integer

If OptionReaming.Value = True Then StrPreHole = "铰"
If OptionCounterboring.Value = True Then StrPreHole = "扩"
If OptionDrilling.Value = True Then StrPreHole = "钻"
If OptionBoring.Value = True Then StrPreHole = "镗"

If (IsNumeric(TextD.Text)) And (IsNumeric(TextL0.Text)) Then
    Recordset1.Open "SELECT * FROM 6_8圆孔拉削余量 WHERE D1<=" & TextD.Text & " AND " & TextD.Text & "<=D2 AND 预制孔加工方式='" & StrPreHole & "'"
    If Recordset1.RecordCount > 0 Then '有记录
        '取得A0-------------------------------------------
        n = 0
        For i = 0 To CInt(Recordset1.Fields.Count) - 1 '遍历字段名
          If (IsNumeric(Recordset1.Fields(i).Name) And Recordset1.Fields(i) > 0) Then '如果字段名为数字且该字段列有内容
            ly(n) = Recordset1.Fields(i).Name
            lydata(n) = Recordset1.Fields(i)
            n = n + 1
          End If
        Next i
        lymin = ly(0)
        lymax = ly(n - 1)
        If Val(TextL0.Text) < lymin Then  'L0低于下限
          GetA0FromTable = "拉削长度L0 < [圆孔拉削余量]表最小查询值 = " & lymin & "！"
          Else
            If Val(TextL0.Text) > lymax Then 'L0高出上限
                GetA0FromTable = "拉削长度L0 > [圆孔拉削余量]表最大查询值 = " & lymax & "！"
            Else 'L0满足范围
            For i = 0 To n - 1
             If Val(TextL0.Text) >= ly(i) Then
                GetA0FromTable = lydata(i) '得到A0
             End If
            Next i
            End If
        End If
    Else '没记录
    GetA0FromTable = "欲符合查表条件请更改预制孔加工方式。"
    End If
    Recordset1.Close
Else
    GetA0FromTable = "请正确设置拉削直径D和拉削长度L0。"
End If
End Function


Private Sub Checkbalpha1_Click()
If Checkbalpha1.Value = 1 Then
    Labelbalpha1_1.Enabled = True
    Labelbalpha1_2.Enabled = True
    Textbalpha1_1.Enabled = True
    Textbalpha1_2.Enabled = True
    Textbalpha1_1.Text = "0.05"
    Textbalpha1_2.Text = "0.05"
Else
    Labelbalpha1_1.Enabled = False
    Labelbalpha1_2.Enabled = False
    Textbalpha1_1.Enabled = False
    Textbalpha1_2.Enabled = False
    Textbalpha1_1.Text = ""
    Textbalpha1_2.Text = ""
End If
End Sub

Private Sub CheckHasChipDividingGroove_Click()
Dim i%
If CheckHasChipDividingGroove.Value = 1 Then
    CommandChipDividingGroove.Enabled = True
    For i = Labelchip.LBound To Labelchip.UBound
        Labelchip(i).Enabled = True
    Next i
    Textnk.Enabled = True
    Textbc.Enabled = True
    Texthc.Enabled = True
    Textrc.Enabled = True
    TextOmegac.Enabled = True
    LabelbcRange.Enabled = True
    LabelhcRange.Enabled = True
    LabelrcRange.Enabled = True
    LabelOmegacRange.Enabled = True
Else
    CommandChipDividingGroove.Enabled = False
    For i = Labelchip.LBound To Labelchip.UBound
        Labelchip(i).Enabled = False
    Next i
    Textnk.Enabled = False
    Textbc.Enabled = False
    Texthc.Enabled = False
    Textrc.Enabled = False
    TextOmegac.Enabled = False
    LabelbcRange.Enabled = False
    LabelhcRange.Enabled = False
    LabelrcRange.Enabled = False
    LabelOmegacRange.Enabled = False
End If
End Sub

Public Sub ComboD1ToleranceZone_GotFocus()
Dim U, L As Double
If IsNumeric(TextD1.Text) Then
    GetLimitFromTable Val(TextD1.Text), ComboD1ToleranceZone.Text, U, L
    TextD1UpperLimit.Text = FixLimit(U / 1000)
    TextD1LowerLimit.Text = FixLimit(L / 1000)
End If
End Sub

Public Sub ComboD2ToleranceZone_GotFocus()
Dim U, L As Double
If IsNumeric(TextD2.Text) Then
    GetLimitFromTable Val(TextD2.Text), ComboD2ToleranceZone.Text, U, L
    TextD2UpperLimit.Text = FixLimit(U / 1000)
    TextD2LowerLimit.Text = FixLimit(L / 1000)
End If
End Sub

Public Sub ComboD3ToleranceZone_GotFocus()
Dim U, L As Double
If IsNumeric(TextD3.Text) Then
    GetLimitFromTable Val(TextD3.Text), ComboD3ToleranceZone.Text, U, L
    TextD3UpperLimit.Text = FixLimit(U / 1000)
    TextD3LowerLimit.Text = FixLimit(L / 1000)
End If
End Sub

Public Sub ComboD4ToleranceZone_GotFocus()
Dim U, L As Double
If IsNumeric(TextD4.Text) Then
    GetLimitFromTable Val(TextD1.Text), ComboD4ToleranceZone.Text, U, L
    TextD4UpperLimit.Text = FixLimit(U / 1000)
    TextD4LowerLimit.Text = FixLimit(L / 1000)
End If
End Sub

Public Sub ComboDToleranceZone_Click()
Dim U#, L#
If IsNumeric(TextD.Text) Then
    GetLimitFromTable Val(TextD.Text), ComboDToleranceZone.Text, U, L
    TextDMax.Text = FixLimit(U / 1000)
    TextDMin.Text = FixLimit(L / 1000)
End If
End Sub

Public Function Checkingh_hz(Optional ByVal ShowMsg As Boolean = True) As Boolean
If Val(Comboh.Text) < Val(LabelhminFormula.Caption) Then '不通过
    Comboh.ForeColor = RGB(255, 0, 0)
    Labelh.ForeColor = RGB(255, 0, 0)
    Checkingh_hz = False
    If ShowMsg = True Then
        TabStrip1.TabIndex = 3
        'MsgBox "粗切齿、过渡齿容屑槽深度h=" & Comboh.Text & "<深度下限hmin=" & LabelhminFormula.Caption
        SendMsgStr "粗切齿、过渡齿容屑槽深度h=" & Comboh.Text & "<深度下限hmin=" & LabelhminFormula.Caption
    End If
Else '通过
    Comboh.ForeColor = RGB(0, 0, 0)
    Labelh.ForeColor = RGB(0, 0, 0)
    Checkingh_hz = True
End If

If Val(Combohz.Text) < Val(LabelhzminFormula.Caption) Then '不通过
    Combohz.ForeColor = RGB(255, 0, 0)
    Labelhz.ForeColor = RGB(255, 0, 0)
    Checkingh_hz = False
    If ShowMsg = True Then
        TabStrip1.TabIndex = 3
        'MsgBox "精切齿、校准齿容屑槽深度hz=" & Combohz.Text & "<深度下限hzmin=" & LabelhzminFormula.Caption
        SendMsgStr "精切齿、校准齿容屑槽深度hz=" & Combohz.Text & "<深度下限hzmin=" & LabelhzminFormula.Caption
    End If
Else '通过
    Combohz.ForeColor = RGB(0, 0, 0)
    Labelhz.ForeColor = RGB(0, 0, 0)
    Checkingh_hz = True
End If

End Function

Private Sub Combog_Click()
If Combog.ListCount > 1 Then
        Comboh.ListIndex = Combog.ListIndex
        Combol_r.ListIndex = Combog.ListIndex
        ComboU_R.ListIndex = Combog.ListIndex
End If
End Sub

Private Sub Combogz_Click()
If Combogz.ListCount > 1 Then
    Combohz.ListIndex = Combogz.ListIndex
    Combol_rz.ListIndex = Combogz.ListIndex
    ComboU_Rz.ListIndex = Combogz.ListIndex
End If
End Sub

Private Sub Comboh_Change()
Checkingh_hz False
End Sub

Private Sub Comboh_Click()
If Comboh.ListCount > 1 Then
    Combog.ListIndex = Comboh.ListIndex
    Combol_r.ListIndex = Comboh.ListIndex
    ComboU_R.ListIndex = Comboh.ListIndex
End If
End Sub

Private Sub Combohz_Change()
Checkingh_hz False
End Sub

Private Sub Combohz_Click()
If Combohz.ListCount > 1 Then
    Combogz.ListIndex = Combohz.ListIndex
    Combol_rz.ListIndex = Combohz.ListIndex
    ComboU_Rz.ListIndex = Combohz.ListIndex
End If
End Sub

Private Sub Combol_r_Click()
If Combol_r.ListCount > 1 Then
    Comboh.ListIndex = Combol_r.ListIndex
    Combog.ListIndex = Combol_r.ListIndex
    ComboU_R.ListIndex = Combol_r.ListIndex
End If
End Sub

Private Sub Combol_rz_Click()
If Combol_rz.ListCount > 1 Then
    Combohz.ListIndex = Combol_rz.ListIndex
    Combogz.ListIndex = Combol_rz.ListIndex
    ComboU_Rz.ListIndex = Combol_rz.ListIndex
End If
End Sub

Sub RefreshQ_l_l0(Optional ByVal iSendMsg As Boolean, Optional ByVal iReason As Integer) '原因：1更改拉床型号；2更改系数
Dim sReason As String
If iReason = 1 Then sReason = "拉床型号"
If iReason = 2 Then sReason = "拉床允许拉力系数"
Recordset1.Open "SELECT * FROM 常用拉床的主要规格 WHERE 拉床型号='" & ComboModel.Text & "'", Connection1, 1, 1
If IsNumeric(TextQCoefficient.Text) And (LabelQ.Caption <> Str(Val(Recordset1.Fields("公称拉力")) * Val(TextQCoefficient.Text))) Then
    LabelQ.Caption = Val(Recordset1.Fields("公称拉力")) * Val(TextQCoefficient.Text)
    If iSendMsg Then SendMsgStr "由于更改了" & sReason & "，拉床允许拉力[Q]已更改为" & LabelQ.Caption
End If
If Textl_l0.Text <> Recordset1.Fields("l0") Then
    Textl_l0.Text = Recordset1.Fields("l0")
    If iSendMsg Then SendMsgStr "由于更改了" & sReason & "，颈部l0已更改为" & Textl_l0.Text
End If
Recordset1.Close
End Sub

Public Sub ComboModel_Click()
RefreshQ_l_l0 True, 1
End Sub

Public Sub ComboToolMaterial_Click()
'读入硬度数据------------------------------------------------------
Recordset1.Open "SELECT * FROM 材料表 WHERE 牌号='" & ComboToolMaterial.Text & "'", Connection1, 1, 1

If (Recordset1.Fields("HBmin") <> 0) And (Recordset1.Fields("HBmax") <> 0) Then
    LabelToolHBRange.Caption = Recordset1.Fields("HBmin") & "~" & Recordset1.Fields("HBmax")
    TextToolHB.Text = Recordset1.Fields("HBmax")
Else
    LabelToolHBRange.Caption = "无数据。"
    TextToolHB.Text = "无数据。"
End If
Recordset1.Close

'读入许用应力数据--------------------------------------------------
Recordset1.Open "SELECT * FROM 材料表 WHERE 牌号='" & ComboToolMaterial.Text & "'"
If StrComp(Recordset1.Fields("6_55类别"), "高速钢") = 0 Then
    TextToolSigmamax1.Text = "350"
    TextToolSigmamax2.Text = "400"
End If
If StrComp(Recordset1.Fields("6_55类别"), "合金钢") = 0 Then
    TextToolSigmamax1.Text = "250"
    TextToolSigmamax2.Text = "300"
End If
If Not (Recordset1.Fields("6_55类别") <> 0) Then '无数据
    TextToolSigmamax1.Text = "无数据。"
    TextToolSigmamax2.Text = "无数据。"
End If
LabelToolSigmamaxRange.Caption = TextToolSigmamax1.Text & "~" & TextToolSigmamax2.Text & "MPa"
Recordset1.Close
End Sub

Private Sub ComboU_R_Click()
If ComboU_R.ListCount > 1 Then
    Comboh.ListIndex = ComboU_R.ListIndex
    Combog.ListIndex = ComboU_R.ListIndex
    Combol_r.ListIndex = ComboU_R.ListIndex
End If
End Sub

Private Sub ComboU_Rz_Click()
If ComboU_Rz.ListCount > 1 Then
    Combohz.ListIndex = ComboU_Rz.ListIndex
    Combogz.ListIndex = ComboU_Rz.ListIndex
    Combol_rz.ListIndex = ComboU_Rz.ListIndex
End If
End Sub

Public Sub ComboWorkpieceMaterial_Click() '刷新工件材料
Recordset1.Open "SELECT * FROM 材料表 WHERE 牌号=" & "'" & ComboWorkpieceMaterial.Text & "'"

If (Recordset1.Fields("抗拉强度下限") <> 0) And (Recordset1.Fields("抗拉强度上限") <> 0) Then '获得抗拉强度下限
    LabelWpSigmabRange.Caption = Recordset1.Fields("抗拉强度下限") & "~" & Recordset1.Fields("抗拉强度上限")
    TextWpSigmab.Text = (Recordset1.Fields("抗拉强度上限") + Recordset1.Fields("抗拉强度下限")) / 2
Else
    LabelWpSigmabRange.Caption = "无数据。"
    TextWpSigmab.Text = "无数据。请自行测定。"
End If

TextGammaoDeg.Text = Recordset1.Fields("6_18切削齿前角")

Recordset1.Close
End Sub

Private Sub Command1_Click()
MsgBox Format(0.152587, "0.####")
'Picture1(0).BackColor = Picture1(0).BackColor + 1
'Command1.Caption = Val(Command1.Caption) + 1
End Sub

Public Sub CommandCalcFinishingTeethD_Click()
Dim Dmax#, DZone#
DZone = Val(TextDMax.Text) - Val(TextDMin.Text)
Dmax = Val(TextD.Text) + Val(TextDMax.Text)
Recordset1.Open "SELECT * FROM 1_1_24拉削时孔的扩张量 WHERE 孔的直径公差1<=" & DZone & " AND " & DZone & "<=孔的直径公差2"
TextFinishingTeethD.Text = Dmax + Recordset1.Fields("扩张量")
LabelFinishingTeethDDelta.Caption = "=" & Dmax & "+" & Fix0(Recordset1.Fields("扩张量"))
Recordset1.Close
End Sub

Public Sub CommandCalcLength_Click()
TextToolLength.Text = Val(TextL1.Text) + Val(TextL2.Text) + Val(Textl_l0.Text) + Val(Textl_l3.Text) + _
    Val(Textl_l.Text) + Val(Textlg.Text) + Val(Textlz.Text) + Val(Textl_l4.Text)
LabelCalcLength.Caption = "= " & TextL1.Text & " + " & TextL2.Text & " + " & Textl_l0.Text & " + " & Textl_l3.Text & " + " & _
    Textl_l.Text & " + " & Textlg.Text & " + " & Textlz.Text & " + " & Textl_l4.Text
End Sub

Public Sub CommandCalcN_Click()
Dim D As Double
Dim n1, n2, n21, n22, n3, n4 As Integer
Dim A, af As Double

If IsNumeric(TextD.Text) Then
    D = Val(TextD.Text)
    A = ValF(TextA0.Text, "拉削余量A", 1)
    af = ValF(Textaf.Text, "齿升量af", 1)
    If GetToleranceGrade(ComboDToleranceZone.Text) <= 8 Then
        Select Case af '设定n2.caption
        Case Is <= 0.15
            Labeln2.Caption = "3~5": n21 = 3: n22 = 5: n2 = 3
        Case Is <= 0.3
            Labeln2.Caption = "5~7": n21 = 5: n22 = 7: n2 = 5
        Case Is > 0.3
            Labeln2.Caption = "6~8": n21 = 6: n22 = 8: n2 = 6
        End Select
        Labeln3.Caption = "4~7": n3 = 4 '设定n3.caption
        Labeln4.Caption = "5~7": n4 = 5 '设定n4.caption
    Else
        If af <= 0.2 Then '设定n2.caption
            Labeln2.Caption = "2~3": n21 = 2: n22 = 3: n2 = 2
        Else
            Labeln2.Caption = "3~5": n21 = 3: n22 = 5: n2 = 3
        End If
        Labeln3.Caption = "2~5": n3 = 2 '设定n3.caption
        Labeln4.Caption = "4~5": n4 = 4 '设定n4.caption
    End If
    Labeln1.Caption = Int(A / (2 * af) + 2 - n21) & "~" & Int(A / (2 * af) + 5 - n22) '设定n1.caption
    n1 = Int(A / (2 * af) + 2 - n21)
    
    Textn1.Text = n1
    Textn2.Text = n2
    Textn3.Text = n3
    Textn4.Text = n4
End If
End Sub

Public Sub CommandChipDividingGroove_Click()
Recordset1.Open "SELECT * FROM 1_1_19拉刀的分屑槽尺寸 WHERE D1<=" & TextD.Text & " AND " & TextD.Text & "<=D2", Connection1, 1, 1
LabelbcRange.Caption = Fix0(Recordset1.Fields("bc1")) & "~" & Fix0(Recordset1.Fields("bc2"))
LabelhcRange.Caption = Fix0(Recordset1.Fields("hc1")) & "~" & Fix0(Recordset1.Fields("hc2"))
LabelrcRange.Caption = Fix0(Recordset1.Fields("rc1")) & "~" & Fix0(Recordset1.Fields("rc2"))
LabelOmegacRange.Caption = "45°~60°"
Textnk.Text = Fix0(Recordset1.Fields("nk"))
Textbc.Text = Fix0(Recordset1.Fields("bc1"))
Texthc.Text = Fix0(Recordset1.Fields("hc1"))
Textrc.Text = Fix0(Recordset1.Fields("rc1"))
TextOmegac.Text = "45"
Recordset1.Close
End Sub

Public Sub InsertTooth(NumClass As Integer)
Dim n%, i%, nCount%
If ListViewTeeth.ListItems.Count > 0 Then
    n = ListViewTeeth.SelectedItem.Index
    nCount = ListViewTeeth.ListItems.Count + 1
    ListViewTeeth.ListItems.Add nCount
    ListViewTeeth.ListItems(nCount).Text = nCount
    
    For i = ListViewTeeth.ListItems.Count To n + 1 Step -1
    'MsgBox i
        ListViewTeeth.ListItems(i).SubItems(1) = ListViewTeeth.ListItems(i - 1).SubItems(1)
        ListViewTeeth.ListItems(i).SubItems(2) = ListViewTeeth.ListItems(i - 1).SubItems(2)
        ListViewTeeth.ListItems(i).SubItems(3) = ListViewTeeth.ListItems(i - 1).SubItems(3)
    Next i
    ListViewTeeth.ListItems(n).Text = n
Else
    ListViewTeeth.ListItems.Add 1
    ListViewTeeth.ListItems.Item(1).Text = "1"
    n = 1
End If

Select Case NumClass
Case Is = 1:
    ListViewTeeth.ListItems(n).SubItems(1) = "粗切齿"
    ListViewTeeth.ListItems(n).SubItems(2) = ""
    ListViewTeeth.ListItems(n).SubItems(3) = "±0.005"
    Textn1.Text = Val(Textn1.Text) + 1
Case Is = 2:
    ListViewTeeth.ListItems(n).SubItems(1) = "过渡齿"
    ListViewTeeth.ListItems(n).SubItems(2) = ""
    ListViewTeeth.ListItems(n).SubItems(3) = "±0.005"
    Textn2.Text = Val(Textn2.Text) + 1
Case Is = 3:
    ListViewTeeth.ListItems(n).SubItems(1) = "精切齿"
    ListViewTeeth.ListItems(n).SubItems(2) = ""
    ListViewTeeth.ListItems(n).SubItems(3) = "-0.005"
    Textn3.Text = Val(Textn3.Text) + 1
Case Is = 4:
    ListViewTeeth.ListItems(n).SubItems(1) = "校准齿"
    ListViewTeeth.ListItems(n).SubItems(2) = ""
    ListViewTeeth.ListItems(n).SubItems(3) = "-0.005"
    Textn4.Text = Val(Textn4.Text) + 1
End Select
End Sub

Public Sub DeleteOneTooth()
Dim n%, i%
If ListViewTeeth.ListItems.Count > 0 Then
    n = ListViewTeeth.SelectedItem.Index
    Select Case ListViewTeeth.ListItems(n).SubItems(1)
    Case Is = "粗切齿":
        Textn1.Text = Val(Textn1.Text) - 1
    Case Is = "过渡齿":
        Textn2.Text = Val(Textn2.Text) - 1
    Case Is = "精切齿":
        Textn3.Text = Val(Textn3.Text) - 1
    Case Is = "校准齿":
        Textn4.Text = Val(Textn4.Text) - 1
    End Select
    ListViewTeeth.ListItems.Remove (n)
    For i = n To ListViewTeeth.ListItems.Count
        ListViewTeeth.ListItems(i).Text = i
    Next i
End If
End Sub

Private Sub CommandDeleteOneTooth_Click()
DeleteOneTooth
End Sub

Private Sub CommandEditTooth_Click()
EditTooth
End Sub

Private Sub CommandInsertTooth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuDelete.Visible = False
mnuEditTooth.Visible = False
Form1.PopupMenu mnuTooth
mnuDelete.Visible = True
mnuEditTooth.Visible = True
End Sub

Private Sub CommandTooth_Click() '设计刀齿直径
Dim D, Dn4Final, Dnow#, Dn1Final#, Dn2Final#
Dim n1, n2, n21, n22, n3, n4 As Integer
Dim A, af, afnow As Double
Dim Sn2#, dd2#, Sn3#, dd3#
Dim i%, j#

If IsNumeric(TextD.Text) Then
    D = Val(TextD.Text)
    Dn4Final = Val(TextFinishingTeethD.Text)
    
    A = ValF(TextA0.Text, "拉削余量A", 1)
    af = ValF(Textaf.Text, "齿升量af", 1)
    
    n1 = Val(Textn1.Text)
    If n1 <= 1 Then n1 = 1
    n2 = Val(Textn2.Text)
    n3 = Val(Textn3.Text)
    n4 = Val(Textn4.Text)
    
    afnow = af * 2
    
    '开始计算各齿直径-----------------------------------------------------------------------------
    ListViewTeeth.ListItems.Clear
    For i = 1 To n1
        ListViewTeeth.ListItems.Add , , i
        ListViewTeeth.ListItems(i).SubItems(1) = "粗切齿"
        Dnow = D - A + i * afnow
        ListViewTeeth.ListItems(i).SubItems(2) = Dnow
        ListViewTeeth.ListItems(i).SubItems(3) = "±0.005"
    Next i
    
    If n2 <> 0 Then
        Dn1Final = Dnow
        Sn2 = D - Dn1Final
        dd2 = Sn2 / AriSquSum(1, n2)
        j = n2 - 1
        For i = n1 + 1 To n1 + n2
            ListViewTeeth.ListItems.Add , , i
            ListViewTeeth.ListItems(i).SubItems(1) = "过渡齿"
            If Dnow + UpTo5(j * dd2) < D Then Dnow = Dnow + UpTo5(j * dd2)
            ListViewTeeth.ListItems(i).SubItems(2) = Dnow
            ListViewTeeth.ListItems(i).SubItems(3) = "±0.005"
            j = j - 1
        Next i
    End If
    
    Dn2Final = Dnow
    Sn3 = Dn4Final - Dn2Final
    'dd3 = Sn3 / AriSquSum(1, n3) '递加算法
    dd3 = Sn3 / (n3 + 1) '平均算法
    j = 0
    
    For i = n1 + n2 + 1 To n1 + n2 + n3 - 1
        'j = j + 1'递加算法
        ListViewTeeth.ListItems.Add , , i
        ListViewTeeth.ListItems(i).SubItems(1) = "精切齿"
        'If Dnow + UpTo5(j * dd3) <= Dn4Final Then Dnow = Dnow + UpTo5(j * dd3) '递加算法
        If Dnow + UpTo5(dd3) <= Dn4Final Then Dnow = Dnow + UpTo5(dd3)  '平均算法
        ListViewTeeth.ListItems(i).SubItems(2) = Dnow
        ListViewTeeth.ListItems(i).SubItems(3) = "-0.005"
    Next i
    
    '最后一个精切齿直径应等于校准齿直径
    ListViewTeeth.ListItems.Add , , n1 + n2 + n3
    ListViewTeeth.ListItems(i).SubItems(1) = "精切齿"
    ListViewTeeth.ListItems(i).SubItems(2) = Dn4Final
    ListViewTeeth.ListItems(i).SubItems(3) = "-0.005"
    
    For i = n1 + n2 + n3 + 1 To n1 + n2 + n3 + n4
        ListViewTeeth.ListItems.Add , , i
        ListViewTeeth.ListItems(i).SubItems(1) = "校准齿"
        ListViewTeeth.ListItems(i).SubItems(2) = Dn4Final '最终尺寸
        ListViewTeeth.ListItems(i).SubItems(3) = "-0.005"
    Next i
End If
End Sub

Private Sub CommandBuild_Click()
Form1.Hide
Form2.Show
MakeMeOnTop Form2.hWnd

oCreateRoundBroach

Unload Form2
Form1.Show
End Sub

Public Sub SetCommandA_Click()
End Sub

Public Sub CommandA_Click() '设置拉削余量
LabelA0Formula.Caption = Fix0(Math.Round(GetA0FromFormula, 2))
LabelA0Table.Caption = Fix0(GetA0FromTable)
TextA0 = Fix0(Math.Round(GetA0FromFormula, 1))
LabelPreHoleD.Caption = Val(TextD.Text) - Val(TextA0.Text)
End Sub

Public Sub Commandaf_Click() '设置齿升量
If IsNumeric(TextD.Text) Then
    Recordset1.Open "SELECT * FROM 6_56圆拉刀的齿升量 WHERE D1<=" & TextD.Text & " AND " & TextD.Text & "<D2", Connection1, 1, 1
    LabelafRange.Caption = Fix0(Recordset1.Fields("分层式Min")) & "~" & Fix0(Recordset1.Fields("分层式Max"))
    LabelMRange.Caption = Fix0(20 * (Val(Recordset1.Fields("分层式Min")) - 0.01) + 1.1) & "~" & Fix0(20 * (Val(Recordset1.Fields("分层式Max")) - 0.01) + 1.1)
    Textaf.Text = Fix0(Recordset1.Fields("分层式Max"))
    Recordset1.Close
Else
    LabelafRange.Caption = "请设定拉孔直径D。"
    Textaf.Text = ""
End If
If IsNumeric(Textaf.Text) Then TextM.Text = Fix0(20 * (Val(Textaf.Text) - 0.01) + 1.1)
End Sub

Public Sub CommandKmin_Click() '计算Kmin
Dim Sigmab As Double
Dim af As Double
Dim HasMaterial As Boolean
Dim StrSQL, StrFieldsName As String

Recordset1.Open "SELECT * FROM 材料表 WHERE 牌号=" & "'" & ComboWorkpieceMaterial.Text & "'", Connection1, 1, 1

If StrComp(Recordset1.Fields("6_17类别"), "钢") = 0 Then
    If IsNumeric(TextWpSigmab.Text) Then '通过抗拉强度选择数据
        Sigmab = Val(TextWpSigmab.Text)
        Select Case Sigmab
        Case Is < 400
            StrFieldsName = "钢抗拉强度less400"
            StrSQL = "SELECT af," & StrFieldsName & " FROM 6_17分层式拉刀的容屑系数"
        Case Is <= 700
            StrFieldsName = "钢抗拉强度in400to700"
            StrSQL = "SELECT af," & StrFieldsName & " FROM 6_17分层式拉刀的容屑系数"
        Case Is > 700
            StrFieldsName = "钢抗拉强度more700"
            StrSQL = "SELECT af," & StrFieldsName & " FROM 6_17分层式拉刀的容屑系数"
        End Select
        
    Else
      MsgBox "请设定工件抗拉强度σb测定值。"
    End If
Else
    StrFieldsName = ComboWorkpieceMaterial.Text
    StrSQL = "SELECT af," & StrFieldsName & " FROM 6_17分层式拉刀的容屑系数"
End If
Recordset1.Close

If StrSQL <> 0 Then '有该金属数据
    If IsNumeric(Textaf.Text) Then
        af = Val(Textaf.Text)
        Select Case af
        Case Is < 0.03:
            StrSQL = StrSQL & " WHERE af='aflessp03'"
        Case Is <= 0.07
            StrSQL = StrSQL & " WHERE af='afinp03top07'"
        Case Is > 0.07
            StrSQL = StrSQL & " WHERE af='afmorep07'"
        End Select
        Recordset1.Open StrSQL, Connection1, 1, 1
        TextKmin.Text = Recordset1.Fields(StrFieldsName)
        Recordset1.Close
    Else
        MsgBox "请设定齿升量af。"
    End If
End If
End Sub

Public Sub Commandp_Click() '计算齿距
Dim pFormula#, pzFormula#
pFormula = ValF(TextM, "M值") * Sqr(ValF(TextL0.Text, "拉削长度L0"))
LabelpFormula.Caption = "p计算值:" & Math.Round(pFormula, 2)

Textp.Text = Math.Round(pFormula, 0)

pzFormula = 0.7 * ValF(Textp.Text, "齿距p计算值")
LabelpzFormula.Caption = "pz计算值:" & pzFormula

Textpz.Text = Math.Round(pzFormula, 0)
End Sub

Public Sub CommandCheck_Click() '进行校核
Dim FieldName As String
Dim HB, Fq, Fmax, D2, l_d, l_dmin, Amin, ze, Q, BroachSigma, ToolSigmamax1, ToolSigmamax2 As Double
Dim PassSigma, PassFmax As Boolean
Dim ResultSigma, ResultFmax As String
'计算ze------------------------------------------------------------------------
ze = Fix(ValF(TextL0.Text, "拉削长度L0") / ValF(Textp.Text, "齿距p", 1) + 1)
Labelze.Caption = ze

'刷新Q------------------------------------------------------------------------
ComboModel_Click
Q = ValF(LabelQ.Caption, "拉床允许拉力[Q]")
LabelQ.Caption = Q & "N"

'进行材料分类----------------------------------------------------------------------
Recordset1.Open "SELECT * FROM 材料表 WHERE 牌号='" & ComboToolMaterial.Text & "'"

HB = ValF(TextToolHB.Text, "刀具硬度测定值HB")
If StrComp(Recordset1.Fields("6_55类别"), "碳钢") = 0 Then
    If HB <= 197 Then
        FieldName = "碳钢lessequ197"
    Else
        If HB <= 229 Then
            FieldName = "碳钢in197to229"
        Else
            FieldName = "碳钢more229"
        End If
    End If
End If
If (StrComp(Recordset1.Fields("6_55类别"), "合金钢") = 0) Or (StrComp(Recordset1.Fields("6_55类别"), "高速钢") = 0) Then
    If HB <= 197 Then
        FieldName = "合金钢lessequ197"
    Else
        If HB <= 229 Then
            FieldName = "合金钢in197to229"
        Else
            FieldName = "合金钢more229"
        End If
    End If
End If
If StrComp(Recordset1.Fields("6_55类别"), "灰铸铁") = 0 Then
    If HB <= 180 Then
        FieldName = "灰铸铁lessequ180"
    Else
        FieldName = "灰铸铁more180"
    End If
End If
If StrComp(Recordset1.Fields("6_55类别"), "可锻铸铁") = 0 Then
    FieldName = "可锻铸铁"
End If
Recordset1.Close

'查表6-48得到F'(Fq)并校核-----------------------------------------------------------
Recordset1.Open "SELECT * FROM 6_48拉刀单位长度切削刃上的拉削力 WHERE af=" & ValF(Textaf.Text, "齿升量af", 0)

If Len(FieldName) > 0 Then '材料分类正确
    If Recordset1.RecordCount > 0 Then '由af,刀具材料对应F'数据
        '计算Fq-----------------------------------
        Fq = Recordset1.Fields(FieldName)
        LabelFq.Caption = Fq & "N/mm"
        '计算Fmax------------------------------
        Fmax = 3.33 * Fq * ValF(TextD.Text, "拉孔直径D") * ze
        LabelFmax.Caption = Fmax & "N"

        '判断最小断面--------------------------------------------------------------------------
        l_d = ValF(TextD.Text, "拉孔直径D") - ValF(TextA0.Text, "拉削余量A") - 2 * ValF(Comboh.Text, "容屑槽深度h")
        D2 = ValF(TextD2.Text, "前柄D2")
        If (D2 <> 0) And (D2 < l_d) Then '前柄D2更小
            l_dmin = D2
            Labell_dmin.Caption = "前柄D2=" & D2 & "mm<" & "第一齿槽底直径d=" & l_d & "mm"
        Else
            If (D2 <> l_d) Then '两者不等于-第一齿槽底直径更小
                l_dmin = l_d
                Labell_dmin.Caption = "第一齿槽底直径d=" & l_d & "mm<" & "前柄D2=" & D2 & "mm"
            Else '两者相等
                l_dmin = l_d
                Labell_dmin.Caption = "前柄D2=" & D2 & "mm=" & "第一齿槽底直径d=" & l_d & "mm"
            End If
        End If
        '计算Amin-------------------------------
        Amin = 3.14 * ((l_dmin / 2) ^ 2)
        LabelAmin.Caption = Math.Round(Amin, 2) & " mm^2"
        '计算许用应力------------------------------
        ToolSigmamax1 = ValF(TextToolSigmamax1.Text, "刀具材料许用应力[σ]")
        ToolSigmamax2 = ValF(TextToolSigmamax2.Text, "刀具材料许用应力[σ]")
        LabelToolSigmamaxRange.Caption = ToolSigmamax1 & "~" & ToolSigmamax2 & "MPa"
        '计算拉应力------------------------------
        If Amin <> 0 Then
            BroachSigma = Fmax / Amin
            LabelBroachSigma.Caption = Math.Round(BroachSigma, 2) & "MPa"
        End If
        '校核--------------------------------------------------------------------------------------
        If (Fmax <= Q) Then
            PassFmax = True '通过许用拉力
            LabelCheckF.Caption = "通过"
            LabelCheckF.ForeColor = RGB(0, 0, 0)
        Else
            PassFmax = False '不通过许用拉力
            ResultFmax = "Fmax=" & LabelFmax.Caption & ">[Q]=" & LabelQ.Caption  '不通过原因
            LabelCheckF.Caption = "未通过：" & ResultFmax
            LabelCheckF.ForeColor = RGB(255, 0, 0)
        End If
        
        If (BroachSigma <= ToolSigmamax2) Then
            PassSigma = True '通过许用应力
            LabelCheckSigma.Caption = "通过"
            LabelCheckSigma.ForeColor = RGB(0, 0, 0)
        Else
            PassSigma = False '不通过许用应力
            ResultSigma = "σ=" & LabelBroachSigma.Caption & ">[σ]=" & TextToolSigmamax2.Text '不通过原因
            LabelCheckSigma.Caption = "未通过：" & ResultSigma
            LabelCheckSigma.ForeColor = RGB(255, 0, 0)
        End If
        
        LabelCheckResult = ""
        LabelAdvice.Caption = ""
        LabelCheckResult.ForeColor = RGB(255, 0, 0)
        LabelAdvice.ForeColor = RGB(255, 0, 0)
        If Not PassSigma Then
            LabelCheckResult.Caption = LabelCheckResult.Caption & "应力校核不合格。"
            If (D2 <> 0) And (D2 < l_d) Then '前柄D2更小
                LabelAdvice.Caption = "因前柄D2为制式刀柄，故可通过更改刀具材料提高刀具材料许用应力[σ]。"
            Else
                LabelAdvice.Caption = "可更改槽类型或槽深以更改第一齿槽底直径；或更改刀具材料以提高刀具材料许用应力[σ]。"
            End If
        End If
        
        If Not PassFmax Then
            LabelCheckResult.Caption = LabelCheckResult.Caption & "拉力校核不合格。"
            LabelAdvice.Caption = LabelAdvice.Caption & "减小af；更改刀具材料；增大齿距p；更改拉床型号或拉床允许拉力系数。"
        End If
        
        If Not Checkingh_hz(False) Then
            LabelCheckResult.Caption = LabelCheckResult.Caption & "槽深低于最小槽深hmin。"
            LabelAdvice.Caption = LabelAdvice.Caption & "更改槽类型或槽深。"
        End If
        
        If PassSigma And PassFmax And Checkingh_hz(False) Then
            LabelCheckResult.Caption = "校核合格。"
            LabelAdvice.Caption = "无"
            LabelCheckResult.ForeColor = RGB(0, 0, 0)
            LabelAdvice.ForeColor = RGB(0, 0, 0)
        End If
        
    Else
        LabelBroachSigma.Caption = "无检索结果。请更改af值。"
        LabelCheckSigma.Caption = "无检索结果。请更改af值。"
    End If
Else
    MsgBox "未找到该材料 " & ComboToolMaterial.Text & " 对应分类。"
End If
Recordset1.Close

End Sub

Public Sub CommandGroove_Click() '容屑槽
Dim StrGrooveTableName As String
Dim p, pz As Double
Dim i, n As Integer
Dim HasData As Boolean

'计算hmin
LabelhminFormula.Caption = Math.Round(1.13 * Sqr(ValF(Textaf.Text, "齿升量af") * ValF(TextKmin.Text, "容屑系数Kmin") * ValF(TextL0.Text, "拉削长度L0")), 2)
LabelhzminFormula.Caption = LabelhminFormula.Caption

'区分需要查找的表
If OptionSmallerGroove.Value = True Then
    StrGrooveTableName = "6_23生产中常用的容屑槽尺寸"
Else
    StrGrooveTableName = "6_24曲线和直线齿背容屑槽计算尺寸"
End If

For i = 3 To 0 Step -1 '倒序遍历槽类型
    If (OptionGroove(i).Value = True) And (OptionGroove(i).Enabled = True) Then '若某项被选中且可选
    
        'p--------------------------------------------------------------------
        If IsNumeric(Textp.Text) Then 'p非数字
            '以p和槽类型查表
            p = Val(Textp.Text)
            
            Recordset1.Open "SELECT * FROM " & StrGrooveTableName & " WHERE 槽类型='" _
                & OptionGroove(i).Caption & "'" & " AND p=" & p
            
            Comboh.Clear
            Combog.Clear
            Combol_r.Clear
            ComboU_R.Clear
            
            HasData = (Recordset1.RecordCount > 0) And _
                (Not IsNull(Recordset1.Fields("h"))) And _
                (Not IsNull(Recordset1.Fields("g"))) And _
                (Not IsNull(Recordset1.Fields("l_r"))) And _
                (Not IsNull(Recordset1.Fields("U_R")))
            If HasData Then '如果有记录且有效
                For n = 0 To Recordset1.RecordCount - 1
                    Comboh.AddItem (Recordset1.Fields("h"))
                    Combog.AddItem (Recordset1.Fields("g"))
                    Combol_r.AddItem (Recordset1.Fields("l_r"))
                    ComboU_R.AddItem (Recordset1.Fields("U_R"))
                    Comboh.ListIndex = 0
                    Combog.ListIndex = 0
                    Combol_r.ListIndex = 0
                    ComboU_R.ListIndex = 0
                    Recordset1.MoveNext
                Next n
                If n > 1 Then '如果记录大于1
                    SendMsgStr "粗切齿、过渡齿容屑槽尺寸还有另外" & Comboh.ListCount - 1 & "组数据，点击箭头更换。"
                    Comboh.ToolTipText = "粗切齿、过渡齿容屑槽尺寸还有另外" & Comboh.ListCount - 1 & "组数据，点击箭头更换。"
                    Combog.ToolTipText = "粗切齿、过渡齿容屑槽尺寸还有另外" & Comboh.ListCount - 1 & "组数据，点击箭头更换。"
                    Combol_r.ToolTipText = "粗切齿、过渡齿容屑槽尺寸还有另外" & Comboh.ListCount - 1 & "组数据，点击箭头更换。"
                    ComboU_R.ToolTipText = "粗切齿、过渡齿容屑槽尺寸还有另外" & Comboh.ListCount - 1 & "组数据，点击箭头更换。"
                Else
                    Comboh.ToolTipText = ""
                    Combog.ToolTipText = ""
                    Combol_r.ToolTipText = ""
                    ComboU_R.ToolTipText = ""
                End If
                'Textp.BackColor = &H80000005  '白
                'Textp.ForeColor = &H80000008 '黑
            Else '如果没有记录或数据无效
                MsgBox "未找到 p = " & p & " 对应容屑槽尺寸数据，请重设齿距p、容屑槽规格决策或槽类型。"
                SendMsgStr "未找到 p = " & p & " 对应容屑槽尺寸数据，请重设齿距p、容屑槽规格决策或槽类型。"
            End If
            Recordset1.Close
        Else '如果p非数字
            MsgBox "请设定齿距p。"
        End If
        
        'pz--------------------------------------------------------------------
        If IsNumeric(Textpz.Text) Then 'pz非数字
            '以pz和槽类型查表
            pz = Val(Textpz.Text)
            
            Recordset1.Open "SELECT * FROM " & StrGrooveTableName & " WHERE 槽类型='" _
                & OptionGroove(i).Caption & "'" & " AND p=" & pz
            
            Combohz.Clear
            Combogz.Clear
            Combol_rz.Clear
            ComboU_Rz.Clear
            
            HasData = (Recordset1.RecordCount > 0) And _
                (Not IsNull(Recordset1.Fields("h"))) And _
                (Not IsNull(Recordset1.Fields("g"))) And _
                (Not IsNull(Recordset1.Fields("l_r"))) And _
                (Not IsNull(Recordset1.Fields("U_R")))
            If HasData Then '如果有记录且有效
                For n = 0 To Recordset1.RecordCount - 1
                    Combohz.AddItem (Recordset1.Fields("h"))
                    Combogz.AddItem (Recordset1.Fields("g"))
                    Combol_rz.AddItem (Recordset1.Fields("l_r"))
                    ComboU_Rz.AddItem (Recordset1.Fields("U_R"))
                    Combohz.ListIndex = 0
                    Combogz.ListIndex = 0
                    Combol_rz.ListIndex = 0
                    ComboU_Rz.ListIndex = 0
                    Recordset1.MoveNext
                Next n
                If n > 1 Then '如果记录大于1
                    SendMsgStr "精切齿、校准齿容屑槽尺寸还有另外" & Combohz.ListCount - 1 & "组数据，点击箭头更换。"
                    Combohz.ToolTipText = "精切齿、校准齿容屑槽尺寸还有另外" & Combohz.ListCount - 1 & "组数据，点击箭头更换。"
                    Combogz.ToolTipText = "精切齿、校准齿容屑槽尺寸还有另外" & Combohz.ListCount - 1 & "组数据，点击箭头更换。"
                    Combol_rz.ToolTipText = "精切齿、校准齿容屑槽尺寸还有另外" & Combohz.ListCount - 1 & "组数据，点击箭头更换。"
                    ComboU_Rz.ToolTipText = "精切齿、校准齿容屑槽尺寸还有另外" & Combohz.ListCount - 1 & "组数据，点击箭头更换。"
                Else
                    Combohz.ToolTipText = ""
                    Combogz.ToolTipText = ""
                    Combol_rz.ToolTipText = ""
                    ComboU_Rz.ToolTipText = ""
                End If
                
                'Textpz.BackColor = &H80000005  '白
                'Textpz.ForeColor = &H80000008 '黑
            Else '如果没有记录或数据无效
                MsgBox "未找到 pz = " & pz & " 对应容屑槽尺寸数据，请重设齿距pz、容屑槽规格决策或槽类型。"
                SendMsgStr "未找到 pz = " & pz & " 对应容屑槽尺寸数据，请重设齿距pz、容屑槽规格决策或槽类型。"
            End If
            Recordset1.Close
        Else '如果pz非数字
            MsgBox "请设定齿距pz。"
        End If
        
    End If
Next i
Checkingh_hz
End Sub

Private Sub CommandAutoSmooth_Click() '光滑尺寸
Recordset1.Open "SELECT * FROM 6_36拉刀圆柱形柄部II型型式和基本尺寸 ORDER BY D1" '以D1排序

While Not Recordset1.EOF
    If (ValF(TextD, "拉孔直径D") - ValF(TextA0.Text, "拉削余量")) >= Recordset1.Fields("D1") Then
        TextD1.Text = Recordset1.Fields("D1")
        TextDq1.Text = Recordset1.Fields("Dq1")
        TextD2.Text = Recordset1.Fields("D2")
        TextL1.Text = Recordset1.Fields("L1")
        TextL2.Text = Recordset1.Fields("L2")
        TextU_L3.Text = Recordset1.Fields("L3")
        TextU_L4.Text = Recordset1.Fields("L4")
        TextC.Text = Recordset1.Fields("C")
    End If
    Recordset1.MoveNext
Wend
Recordset1.Close
Call ComboD1ToleranceZone_GotFocus
Call ComboD2ToleranceZone_GotFocus
'颈部-----------------------------------------------------------
TextD0.Text = TextD1.Text

'过渡锥l--------------------------------------------------------
Combolq3.Clear
Combolq3.AddItem "10"
Combolq3.AddItem "15"
Combolq3.AddItem "20"
Combolq3.ListIndex = 0

'前导部-------------------------------------------------------------
TextD3.Text = ValF(TextD.Text, "拉孔直径D") - ValF(TextA0.Text, "拉削余量A") 'D3=D0
Call ComboD3ToleranceZone_GotFocus '更新公差
Textl_l3.Text = TextL0.Text

'切削部及校准部-----------------------------------------------------
Textl_l.Text = (ValF(Textn1.Text, "粗切齿齿数") + ValF(Textn2.Text, "过渡齿齿数")) * ValF(Textp.Text, "齿距p")
Textlg.Text = ValF(Textn3.Text, "精切齿齿数") * ValF(Textpz.Text, "齿距pz")
Textlz.Text = ValF(Textn4.Text, "校准齿齿数") * ValF(Textpz.Text, "齿距pz")

'后导部--------------------------------------------------------------
TextD4.Text = TextD.Text
Call ComboD4ToleranceZone_GotFocus '更新公差
Recordset1.Open "SELECT * FROM 6_40工件长度和后导部长度 WHERE 工件长度L1<=" & ValF(TextL0.Text, "拉削长度L") & " AND " & ValF(TextL0.Text, "拉削长度L") & "<工件长度L2", Connection1, 1, 1
Textl_l4.Text = Recordset1.Fields("后导部长度")
Recordset1.Close

'拉刀总长-------------------------------------------------------------
Call CommandCalcLength_Click
End Sub

Public Sub ShowMsgAutoExit(ByVal Msg As String, ByVal Time As Integer)
Form2.Label1.Caption = Msg
Form2.Timer1.Interval = Time
Form2.Timer1.Enabled = True
Form2.Show
End Sub

Private Sub CommandAuto_Click()
SendMsgStr "自动推理开始。"
ShowMsgAutoExit "正在自动推理...", 800

Call CommandA_Click
Call Commandaf_Click
Call CommandKmin_Click
Call Commandp_Click

Call CommandCalcN_Click
Call CommandCalcFinishingTeethD_Click
Call CommandTooth_Click

Call CommandGroove_Click
Call CommandChipDividingGroove_Click

Call CommandAutoSmooth_Click

Call CommandCheck_Click

InitListViewParameters
RefreshListViewParameters

SendMsgStr "已完成自动推理。" & "D=" & TextD.Text & " L=" & TextL0.Text
End Sub

Sub iregsvr32(ByVal FileName As String)
If Dir(FileName) <> "" Then
Else
    Shell ("regsvr32 /s " & FileName)
    SendMsgStr "检测到系统没有" & FileName & "文件，已完成注册。"
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
'进行OCX环境初始化-------------------------------------
iregsvr32 "C:\windows\system32\MSADODC.OCX"
'iregsvr32 "C:\windows\system32\COMCTL32.OCX"
iregsvr32 "C:\windows\system32\comdlg32.OCX"

'初始化TabStrip----------------------------------------
With TabStrip1
    .Tabs.Item(1).Caption = "刀具基本参数"
    .Tabs.Add , , "设计系数选定"
    .Tabs.Add , , "设计刀齿直径"
    .Tabs.Add , , "容屑槽与分屑槽"
    .Tabs.Add , , "光滑部分尺寸"
    .Tabs.Add , , "强度校核"
    .Tabs.Add , , "参数列表"
End With
TabStrip1.Top = 0
TabStrip1.Left = 0
TabStrip1.Height = 9255
TabStrip1.Width = 8535
TabStrip1_Click

'初始化From-----------------------------------------------
'Form1.Height = 11115
Form1.Width = 8535 '10620

'读入数据库---------------------------------------------
DatabaseName = App.Path & "\拉刀设计数据库.mdb"
Connection1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseName
SendMsgStr "已读入" & DatabaseName

'载入刀具类型-----------------------------------------
ComboBroach.AddItem ("分层式")
ComboBroach.AddItem ("综合式(待完成)")
ComboBroach.AddItem ("分块式(待完成)")
ComboBroach.ListIndex = 0

'载入拉床型号及最大允许拉力[Q]-------------------------------
Recordset1.Open "SELECT * FROM 常用拉床的主要规格", Connection1, 1, 1
While Not Recordset1.EOF
    ComboModel.AddItem (Recordset1.Fields("拉床型号"))
    Recordset1.MoveNext
Wend
Recordset1.Close
ComboModel.ListIndex = 0

'设置机床拉力系数---------------------------------------------------
'OptionAutoQCoefficient.Value = True
'OptionNew.Value = True
LabelQCoefficient.Caption = "新机床0.9~1；" & Chr(13) & Chr(10) & _
                            "处于良好状态的旧机床0.8；" & Chr(13) & Chr(10) & _
                            "处于不良状态下的旧机床0.5~0.7。"

'点击铰加工单选框-----------------------------------------
OptionReaming.Value = True

'载入材料------------------------------------------------
Recordset1.Open "SELECT * FROM 材料表", Connection1, 1, 1

While Not Recordset1.EOF '读入刀具及工具材料类型
    ComboToolMaterial.AddItem (Recordset1.Fields("牌号"))
    ComboWorkpieceMaterial.AddItem (Recordset1.Fields("牌号"))
    Recordset1.MoveNext
Wend
Recordset1.Close
For i = 0 To ComboToolMaterial.ListCount - 1 '设置默认刀具材料
    If StrComp(ComboToolMaterial.List(i), "W18Cr4V") = 0 Then
        ComboToolMaterial.ListIndex = i
    End If
Next i

For i = 0 To ComboWorkpieceMaterial.ListCount - 1 '设置默认工具材料
    If StrComp(ComboWorkpieceMaterial.List(i), "45") = 0 Then
        ComboWorkpieceMaterial.ListIndex = i
    End If
Next i

'点击刀具材料复选框以更新许用应力数据----------------------------
Call ComboToolMaterial_Click

'点击工具材料复选框以更新抗拉强度数据----------------------------
Call ComboWorkpieceMaterial_Click

'初始化D数据---------------------------------------------
TextD_Validate (False) '刷新D

'读入拉孔直径精度等级--------------------------------------------
ComboDToleranceZone.AddItem ("H7")
ComboDToleranceZone.AddItem ("H8")
ComboDToleranceZone.AddItem ("H9")
ComboDToleranceZone.ListIndex = 0

'载入光滑部分公差带---------------------------------------------
ComboD1ToleranceZone.AddItem "f8"
ComboD2ToleranceZone.AddItem "h12"
ComboD3ToleranceZone.AddItem "e8"
ComboD4ToleranceZone.AddItem "f7"
ComboD1ToleranceZone.ListIndex = 0
ComboD2ToleranceZone.ListIndex = 0
ComboD3ToleranceZone.ListIndex = 0
ComboD4ToleranceZone.ListIndex = 0

'刀齿直径列表初始化-----------------------------------------------
InitListViewTeeth

'传入参数-----------------------------------------------------
Dim nCmdString As String
nCmdString = Command
nCmdString = Replace(nCmdString, """", "")
If nCmdString <> "" Then
    CommonDialog1.FileName = nCmdString
    MenuOpen_Click
End If

End Sub

Private Sub LabelToolSigmaMax1_Click()

End Sub

Private Sub ListViewTeeth_DblClick()
EditTooth
End Sub

Sub EditTooth()
'Dim alln% '能够显示的行数
Dim NowN% '当前行数
If (ListViewTeeth.ColumnHeaders.Count > 0) And (ListViewTeeth.ListItems.Count > 0) Then
    
    '移动编辑框
    TextEditTooth.Left = ListViewTeeth.Left + (ListViewTeeth.ColumnHeaders.Item(1).Width + ListViewTeeth.ColumnHeaders.Item(2).Width) * 1.65
    
    NowN = Fix(pY / ListViewTeeth.ListItems(ListViewTeeth.SelectedItem.Index).Height)
    
    TextEditTooth.Top = ListViewTeeth.Top + ListViewTeeth.ListItems(ListViewTeeth.SelectedItem.Index).Height * NowN + 80
    TextEditTooth.Width = ListViewTeeth.ColumnHeaders.Item(3).Width
    TextEditTooth.Height = ListViewTeeth.ListItems(ListViewTeeth.SelectedItem.Index).Height
    
    If NowN <= ListViewTeeth.ListItems.Count Then
        TextEditTooth.Text = ListViewTeeth.ListItems(ListViewTeeth.SelectedItem.Index).SubItems(2)
        OnEditToothIndex = ListViewTeeth.SelectedItem.Index
        TextEditTooth.Visible = True
        TextEditTooth.SetFocus
        TextEditTooth.SelStart = 0
        TextEditTooth.SelLength = Len(TextEditTooth.Text)
    End If
End If
End Sub

Private Sub ListViewTeeth_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Or KeyCode = 110 Then DeleteOneTooth
End Sub

Private Sub ListViewTeeth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pX = X
pY = Y
TextEditTooth.Visible = False
If Button = 2 Then
    Form1.PopupMenu mnuTooth
End If
End Sub

Private Sub MenuAbout_Click()
frmAbout.Show
End Sub

Private Sub MenuExit_Click()
Unload Me
End Sub

Private Sub MenuOpen_Click()
'CancelError 为 True。
On Error GoTo ErrHandler
CommonDialog1.Filter = "所有文件 (*.*)|*.*|圆孔拉刀数据文件 (*.DAT)|*.DAT" '设置过滤器。
CommonDialog1.FilterIndex = 2 '指定缺省过滤器。
CommonDialog1.ShowOpen '显示“保存”对话框。
OpenBroach CommonDialog1.FileName '调用打开文件的过程。
ShowMsgAutoExit "打开完毕。", 800
SendMsgStr "文件" & CommonDialog1.FileName & "打开完毕。"

Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub MenuSave_Click()
'MsgBox Dir(CommonDialog1.FileName)
If CommonDialog1.FileName <> "" Then '文件存在
    SaveBroach CommonDialog1.FileName '调用打开文件的过程。
    ShowMsgAutoExit "保存完毕。", 800
    SendMsgStr "文件" & CommonDialog1.FileName & "保存完毕。"
Else '文件不存在
    MenuSaveAs_Click
End If
End Sub

Private Sub MenuSaveAs_Click()
'CancelError 为 True。
On Error GoTo ErrHandler
CommonDialog1.Filter = "所有文件 (*.*)|*.*|圆孔拉刀数据文件 (*.DAT)|*.DAT" '设置过滤器。
CommonDialog1.FilterIndex = 2 '指定缺省过滤器。
CommonDialog1.ShowSave '显示“保存”对话框。
SaveBroach CommonDialog1.FileName '调用保存文件的过程。
ShowMsgAutoExit "保存完毕。", 800
SendMsgStr "文件" & CommonDialog1.FileName & "保存完毕。"

Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub mnuDelete_Click()
DeleteOneTooth
End Sub

Private Sub mnuEditTooth_Click()
EditTooth
End Sub

Private Sub mnuInsertN1_Click()
InsertTooth (1)
End Sub
Private Sub mnuInsertN2_Click()
InsertTooth (2)
End Sub
Private Sub mnuInsertN3_Click()
InsertTooth (3)
End Sub
Private Sub mnuInsertN4_Click()
InsertTooth (4)
End Sub

Sub CalcTeethLength()
Textl_l.Text = (Val(Textn1.Text) + Val(Textn2.Text)) * Val(Textp.Text)
Textlg.Text = Val(Textn3.Text) * Val(Textpz.Text)
Textlz.Text = Val(Textn4.Text) * Val(Textpz.Text)
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer
For i = 0 To 6 'TabStrip1.Tabs.Count
    Picture1(i).Visible = False
Next i
Picture1(TabStrip1.SelectedItem.Index - 1).BackColor = &H8000000F
Picture1(TabStrip1.SelectedItem.Index - 1).BorderStyle = 0
Picture1(TabStrip1.SelectedItem.Index - 1).Top = 360
Picture1(TabStrip1.SelectedItem.Index - 1).Left = (TabStrip1.Width - Picture1(TabStrip1.SelectedItem.Index - 1).Width) / 2
Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True

If TabStrip1.SelectedItem.Index - 1 = 4 Then
    CalcTeethLength
End If

If TabStrip1.SelectedItem.Index - 1 = 6 Then
    InitListViewParameters
    RefreshListViewParameters
End If
End Sub

Private Sub TextA0_Change()
If IsNumeric(TextD.Text) And IsNumeric(TextA0.Text) Then
    LabelPreHoleD.Caption = Val(TextD.Text) - Val(TextA0.Text)
    SendMsgStr "拉削余量A已改变。"
Else
    LabelPreHoleD.Caption = "请检查拉孔直径D与拉削余量A值。"
    SendMsgStr "请检查拉孔直径D与拉削余量A值。"
End If
End Sub

Private Sub TextAutoQCoefficient_Change()

End Sub

Private Sub TextD_Validate(Cancel As Boolean) '完成编辑D
Dim D, dymin, dymax As Integer
D = ValF(TextD.Text, "拉孔直径D")
'取得D允许范围----------------------------------------
Recordset1.Open "SELECT * FROM 6_8圆孔拉削余量 ORDER BY D1", Connection1, 1, 1
dymin = Recordset1.Fields("D1")
Recordset1.MoveLast
dymax = Recordset1.Fields("D2")
Recordset1.Close
If D < dymin Then MsgBox "拉孔直径D < [圆孔拉削余量]表最小查询值 = " & dymin & "！"
If D > dymax Then MsgBox "拉孔直径D > [圆孔拉削余量]表最大查询值 = " & dymax & "！"

If ComboDToleranceZone.ListCount > 0 Then
    Call ComboDToleranceZone_Click
End If
End Sub

Private Sub TextEditTooth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TextEditTooth_LostFocus
End If
End Sub

Private Sub TextEditTooth_LostFocus()
If ListViewTeeth.ListItems.Count > 0 Then
    ListViewTeeth.ListItems(OnEditToothIndex).SubItems(2) = TextEditTooth.Text
End If
TextEditTooth.Visible = False
End Sub

Private Sub TextManualQCoefficient_Validate(Cancel As Boolean)
Call ComboModel_Click
End Sub

Private Sub TextQCoefficient_Change()
RefreshQ_l_l0 True, 2
End Sub
