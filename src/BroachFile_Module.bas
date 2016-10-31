Attribute VB_Name = "BroachFile_Module"
Option Explicit

Option Base 1 '数组下界为1

Type BroachVars
    ParametersName As String '参数名称
    VarsName As String '变量
    Value As String '变量值
End Type
Const BCount% = 76
Dim BroachData(BCount) As BroachVars
Dim BroachDataCount As Integer
Dim BroachDataIndex As Integer

Type BroachToothVars
    ToothNum As String
    ToothClass As String
    ToothD As String
    ToothLimit As String
End Type
Dim BroachTeethData() As BroachToothVars
Dim BroachTeethDataCount As Integer

'得到一个参数
Sub AddOneBroachVar(ByVal sParametersName As String, Optional ByVal sVarsName As String, Optional ByVal sValue As String)
BroachDataCount = BroachDataCount + 1
BroachData(BroachDataCount).ParametersName = sParametersName
BroachData(BroachDataCount).VarsName = sVarsName
BroachData(BroachDataCount).Value = sValue
End Sub

Function ReadOneBroachVar() As String '置入一个控件
BroachDataIndex = BroachDataIndex + 1
ReadOneBroachVar = BroachData(BroachDataIndex).Value
End Function

Sub Tab1ToRecord() '得到参数-共14条
With Form1
    AddOneBroachVar "实例编号", , .TextBroachNum.Text
    AddOneBroachVar "实例名称", , .TextBroachName.Text
    AddOneBroachVar "刀具类型", , .ComboBroach.Text
    AddOneBroachVar "拉孔直径", "D", .TextD.Text
    AddOneBroachVar "拉孔直径公差", "", .ComboDToleranceZone.Text
    AddOneBroachVar "拉削长度", "L0", .TextL0.Text
    AddOneBroachVar "预制孔加工方式", , .GetStrFromMachiningMode
    AddOneBroachVar "拉床型号", , .ComboModel.Text
    AddOneBroachVar "刀具材料", , .ComboToolMaterial.Text
    AddOneBroachVar "刀具硬度测定值", "HB", .TextToolHB.Text
    AddOneBroachVar "刀具材料许用应力下界", "[σ]1", .TextToolSigmamax1.Text
    AddOneBroachVar "刀具材料许用应力上界", "[σ]2", .TextToolSigmamax2.Text
    AddOneBroachVar "工件材料", , .ComboWorkpieceMaterial.Text
    AddOneBroachVar "工件抗拉强度测定值", "σb", .TextWpSigmab.Text
End With
End Sub

Sub RecordToTab1() '置入控件
With Form1
    .TextBroachNum.Text = ReadOneBroachVar '"实例编号"
    .TextBroachName.Text = ReadOneBroachVar '"实例名称"
    .SetComboFromStr .ComboBroach, ReadOneBroachVar '"刀具类型"
    .TextD.Text = ReadOneBroachVar '"拉孔直径"
    .SetComboFromStr .ComboDToleranceZone, ReadOneBroachVar: .ComboDToleranceZone_Click '"拉孔直径公差"
    .TextL0.Text = ReadOneBroachVar '"拉削长度"
    .SetMachiningModeFromStr (ReadOneBroachVar) '"预制孔加工方式"
    .SetComboFromStr .ComboModel, ReadOneBroachVar: .ComboModel_Click '"拉床型号"
    .SetComboFromStr .ComboToolMaterial, ReadOneBroachVar: .ComboToolMaterial_Click '"刀具材料"
    .TextToolHB.Text = ReadOneBroachVar '"刀具硬度测定值"
    .TextToolSigmamax1.Text = ReadOneBroachVar '"刀具材料许用应力下界"
    .TextToolSigmamax2.Text = ReadOneBroachVar '"刀具材料许用应力上界"
    .SetComboFromStr .ComboWorkpieceMaterial, ReadOneBroachVar: .ComboWorkpieceMaterial_Click '"工件材料"
    .TextWpSigmab.Text = ReadOneBroachVar '"工件抗拉强度测定值"
End With
End Sub

Sub Tab2ToRecord() '共18条
With Form1
    AddOneBroachVar "拉削余量", "A", .TextA0.Text
    AddOneBroachVar "齿升量", "af", .Textaf.Text
    AddOneBroachVar "M值", "M", .TextM.Text
    AddOneBroachVar "容屑系数", "Kmin", .TextKmin.Text
    AddOneBroachVar "切削齿齿距", "p", .Textp.Text
    AddOneBroachVar "校准齿齿距", "pz", .Textpz.Text
    AddOneBroachVar "是否留刃带", , .Checkbalpha1.Value
    AddOneBroachVar "前角(角度分量)", "γo(Deg)", .TextGammaoDeg.Text
    AddOneBroachVar "前角(分分量)", "γo(Min)", .TextGammaoMin.Text
    AddOneBroachVar "前角(秒分量)", "γo(Sec)", .TextGammaoSec.Text
    AddOneBroachVar "切削齿后角(角度分量)", "αo(Deg)", .TextAlphaoDeg.Text
    AddOneBroachVar "切削齿后角(分分量)", "αo(Min)", .TextAlphaoMin.Text
    AddOneBroachVar "切削齿后角(秒分量)", "αo(Sec)", .TextAlphaoSec.Text
    AddOneBroachVar "切削齿刃带宽", "bα1_1", .Textbalpha1_1.Text
    AddOneBroachVar "校准齿后角(角度分量)", "αoz(Deg)", .TextAlphaozDeg.Text
    AddOneBroachVar "校准齿后角(分分量)", "αoz(Min)", .TextAlphaozMin.Text
    AddOneBroachVar "校准齿后角(秒分量)", "αoz(Sec)", .TextAlphaozSec.Text
    AddOneBroachVar "校准齿刃带宽", "bα1_2", .Textbalpha1_2.Text
End With
End Sub

Sub RecordToTab2()
With Form1
    .CommandA_Click: .TextA0.Text = ReadOneBroachVar '"拉削余量"
    .Commandaf_Click: .Textaf.Text = ReadOneBroachVar '"齿升量"
    .TextM.Text = ReadOneBroachVar '"M值"
    .CommandKmin_Click: .TextKmin.Text = ReadOneBroachVar '"容屑系数"
    .Commandp_Click: .Textp.Text = ReadOneBroachVar '"切削齿齿距"
    .Textpz.Text = ReadOneBroachVar '"校准齿齿距"
    .Checkbalpha1.Value = ReadOneBroachVar '"是否留刃带"
    .TextGammaoDeg.Text = ReadOneBroachVar '"前角(角度分量)"
    .TextGammaoMin.Text = ReadOneBroachVar '"前角(分分量)"
    .TextGammaoSec.Text = ReadOneBroachVar '"前角(秒分量)"
    .TextAlphaoDeg.Text = ReadOneBroachVar '"切削齿后角(角度分量)"
    .TextAlphaoMin.Text = ReadOneBroachVar '"切削齿后角(分分量)"
    .TextAlphaoSec.Text = ReadOneBroachVar '"切削齿后角(秒分量)"
    .Textbalpha1_1.Text = ReadOneBroachVar '"切削齿刃带宽"
    .TextAlphaozDeg.Text = ReadOneBroachVar '"校准齿后角(角度分量)"
    .TextAlphaozMin.Text = ReadOneBroachVar '"校准齿后角(分分量)"
    .TextAlphaozSec.Text = ReadOneBroachVar '"校准齿后角(秒分量)"
    .Textbalpha1_2.Text = ReadOneBroachVar '"校准齿刃带宽"
End With
End Sub

Sub Tab3ToRecord() '共4条
With Form1
    AddOneBroachVar "粗切齿齿数", , .Textn1.Text
    AddOneBroachVar "过渡齿齿数", , .Textn2.Text
    AddOneBroachVar "精切齿齿数", , .Textn3.Text
    AddOneBroachVar "校准齿齿数", , .Textn4.Text
    AddOneBroachVar "校准齿直径", , .TextFinishingTeethD.Text
End With
End Sub

Sub RecordToTab3()
With Form1
    .CommandCalcN_Click: .Textn1.Text = ReadOneBroachVar '"粗切齿齿数"
    .Textn2.Text = ReadOneBroachVar '"过渡齿齿数"
    .Textn3.Text = ReadOneBroachVar '"精切齿齿数"
    .Textn4.Text = ReadOneBroachVar '"校准齿齿数"
    .CommandCalcFinishingTeethD_Click: .TextFinishingTeethD.Text = ReadOneBroachVar
End With
End Sub

Sub Tab4ToRecord()
With Form1
    AddOneBroachVar "容屑槽规格决策", , .GetStrFromSmallerOption
    AddOneBroachVar "槽类型", , .GetStrFromGrooveOption
    AddOneBroachVar "粗切齿、过渡齿容屑槽深度", "h", .Comboh.Text
    AddOneBroachVar "粗切齿、过渡齿容屑槽齿厚", "g", .Combog.Text
    AddOneBroachVar "粗切齿、过渡齿容屑槽槽底圆弧半径", "r", .Combol_r.Text
    AddOneBroachVar "粗切齿、过渡齿容屑槽齿背圆弧半径", "R", .ComboU_R.Text
    AddOneBroachVar "精切齿、校准齿容屑槽深度", "hz", .Combohz.Text
    AddOneBroachVar "精切齿、校准齿容屑槽齿厚", "gz", .Combogz.Text
    AddOneBroachVar "精切齿、校准齿容屑槽槽底圆弧半径", "rz", .Combol_rz.Text
    AddOneBroachVar "精切齿、校准齿容屑槽齿背圆弧半径", "Rz", .ComboU_Rz.Text
    AddOneBroachVar "是否开分屑槽", , .CheckHasChipDividingGroove.Value
    AddOneBroachVar "分屑槽数量", "nk", .Textnk.Text
    AddOneBroachVar "分屑槽宽度", "bc", .Textbc.Text
    AddOneBroachVar "分屑槽深度", "hc", .Texthc.Text
    AddOneBroachVar "分屑槽圆角半径", "rc", .Textrc.Text
    AddOneBroachVar "分屑槽角度", "ωc", .TextOmegac.Text
    AddOneBroachVar "分屑槽后角增量（角度分量）", "Δαc(Deg)", .TextDeltaAlphacDeg.Text
    AddOneBroachVar "分屑槽后角增量（分分量）", "Δαc(Min)", .TextDeltaAlphacMin.Text
    AddOneBroachVar "分屑槽后角增量（秒分量）", "Δαc(Sec)", .TextDeltaAlphacSec.Text
End With
End Sub

Sub RecordToTab4()
With Form1
    .SetSmallerOptionFromStr ReadOneBroachVar '"容屑槽规格决策"
    .SetGrooveOptionFromStr ReadOneBroachVar '"槽类型"
    .CommandGroove_Click: .Comboh.Text = ReadOneBroachVar 'h
    .Combog.Text = ReadOneBroachVar 'g
    .Combol_r.Text = ReadOneBroachVar 'r
    .ComboU_R.Text = ReadOneBroachVar 'R
    .Combohz.Text = ReadOneBroachVar 'hz
    .Combogz.Text = ReadOneBroachVar 'gz
    .Combol_rz.Text = ReadOneBroachVar 'rz
    .ComboU_Rz.Text = ReadOneBroachVar 'Rz
    .CheckHasChipDividingGroove.Value = ReadOneBroachVar '"是否开分屑槽"
    .CommandChipDividingGroove_Click: .Textnk.Text = ReadOneBroachVar
    .Textbc.Text = ReadOneBroachVar
    .Texthc.Text = ReadOneBroachVar
    .Textrc.Text = ReadOneBroachVar
    .TextOmegac.Text = ReadOneBroachVar
    .TextDeltaAlphacDeg.Text = ReadOneBroachVar
    .TextDeltaAlphacMin.Text = ReadOneBroachVar
    .TextDeltaAlphacSec.Text = ReadOneBroachVar
End With
End Sub

Sub Tab5ToRecord()
With Form1
    AddOneBroachVar "前柄D1", "D1", .TextD1.Text
    AddOneBroachVar "前柄L1", "L1", .TextL1.Text
    AddOneBroachVar "前柄D'1", "D'1", .TextDq1.Text
    AddOneBroachVar "前柄D2", "D2", .TextD2.Text
    AddOneBroachVar "前柄L2", "L2", .TextL2.Text
    AddOneBroachVar "柄部参考尺寸L3", "L3", .TextU_L3.Text
    AddOneBroachVar "柄部参考尺寸L4", "L4", .TextU_L4.Text
    AddOneBroachVar "倒角C", "C", .TextC.Text
    AddOneBroachVar "颈部D0", "D0", .TextD0.Text
    AddOneBroachVar "颈部l0", "l0", .Textl_l0.Text
    AddOneBroachVar "过渡锥l'3", "l'3", .Combolq3.Text
    AddOneBroachVar "前导部D3", "D3", .TextD3.Text
    AddOneBroachVar "前导部l3", "l3", .Textl_l3.Text
    AddOneBroachVar "切削部长度l", "l", .Textl_l.Text
    AddOneBroachVar "精切部长度lg", "lg", .Textlg.Text
    AddOneBroachVar "校准部长度lz", "lz", .Textlz.Text
    AddOneBroachVar "后导部D4", "D4", .TextD4.Text
    AddOneBroachVar "后导部l4", "l4", .Textl_l4.Text
    AddOneBroachVar "拉刀总长L", "L", .TextToolLength.Text
End With
End Sub

Sub RecordToTab5()
With Form1
    .TextD1.Text = ReadOneBroachVar '"前柄D1"
    .TextL1.Text = ReadOneBroachVar '"前柄L1"
    .TextDq1.Text = ReadOneBroachVar '"前柄D'1"
    .TextD2.Text = ReadOneBroachVar '"前柄D2"
    .TextL2.Text = ReadOneBroachVar '"前柄L2"
    .TextU_L3.Text = ReadOneBroachVar '"柄部参考尺寸L3"
    .TextU_L4.Text = ReadOneBroachVar '"柄部参考尺寸L4"
    .TextC.Text = ReadOneBroachVar '"倒角C"
    .TextD0.Text = ReadOneBroachVar '
    .Textl_l0.Text = ReadOneBroachVar '
    .Combolq3.Text = ReadOneBroachVar '
    .TextD3.Text = ReadOneBroachVar '
    .Textl_l3.Text = ReadOneBroachVar '
    .Textl_l.Text = ReadOneBroachVar '
    .Textlg.Text = ReadOneBroachVar '
    .Textlz.Text = ReadOneBroachVar '
    .TextD4.Text = ReadOneBroachVar '
    .Textl_l4.Text = ReadOneBroachVar '
    .TextToolLength.Text = ReadOneBroachVar '
    
    .ComboD1ToleranceZone_GotFocus
    .ComboD2ToleranceZone_GotFocus
    .ComboD3ToleranceZone_GotFocus
    .ComboD4ToleranceZone_GotFocus
    .CommandCalcLength_Click
End With
End Sub

Sub Tab6ToRecord()
With Form1
    AddOneBroachVar "拉床允许拉力系数", , .TextQCoefficient.Text
End With
End Sub

Sub RecordToTab6()
With Form1
    .TextQCoefficient.Text = ReadOneBroachVar: .CommandCheck_Click
End With
End Sub

Sub ToothToRecord()
Dim i%
With Form1
    For i = 1 To BroachTeethDataCount
        BroachTeethData(i).ToothNum = .ListViewTeeth.ListItems(i).Text
        BroachTeethData(i).ToothClass = .ListViewTeeth.ListItems(i).SubItems(1)
        BroachTeethData(i).ToothD = .ListViewTeeth.ListItems(i).SubItems(2)
        BroachTeethData(i).ToothLimit = .ListViewTeeth.ListItems(i).SubItems(3)
    Next i
End With
End Sub

Sub RecordToTooth()
Dim i%
With Form1
    .ListViewTeeth.ListItems.Clear
    For i = 1 To BroachTeethDataCount
        .ListViewTeeth.ListItems.Add i
        .ListViewTeeth.ListItems(i).Text = BroachTeethData(i).ToothNum
        .ListViewTeeth.ListItems(i).SubItems(1) = BroachTeethData(i).ToothClass
        .ListViewTeeth.ListItems(i).SubItems(2) = BroachTeethData(i).ToothD
        .ListViewTeeth.ListItems(i).SubItems(3) = BroachTeethData(i).ToothLimit
    Next i
End With
End Sub

Sub RefreshListViewParameters()
Dim i%
BroachDataCount = 0
Tab1ToRecord
Tab2ToRecord
Tab3ToRecord
Tab4ToRecord
Tab5ToRecord
Tab6ToRecord
With Form1
    .ListViewParameters.ListItems.Clear
    For i = 1 To BCount
        .ListViewParameters.ListItems.Add , , i
        .ListViewParameters.ListItems(i).SubItems(1) = BroachData(i).ParametersName
        .ListViewParameters.ListItems(i).SubItems(2) = BroachData(i).VarsName
        .ListViewParameters.ListItems(i).SubItems(3) = BroachData(i).Value
    Next i
End With
End Sub

Sub SaveBroach(ByVal FileName As String) '保存拉刀
Dim i%
BroachDataCount = 0
Tab1ToRecord
Tab2ToRecord
Tab3ToRecord
Tab4ToRecord
Tab5ToRecord
Tab6ToRecord

BroachTeethDataCount = Form1.ListViewTeeth.ListItems.Count
If BroachTeethDataCount > 0 Then
    ReDim BroachTeethData(BroachTeethDataCount) As BroachToothVars
    ToothToRecord
End If

Open FileName For Random As #1 Len = 128
For i = 1 To BCount
    Put #1, , BroachData(i)
Next i

Put #1, , BroachTeethDataCount
If BroachTeethDataCount > 0 Then
    For i = 1 To BroachTeethDataCount
        Put #1, , BroachTeethData(i)
    Next i
End If

Close #1
End Sub

Sub OpenBroach(ByVal FileName As String) '读取拉刀
Dim i%

Open FileName For Random As #1 Len = 128
For i = 1 To BCount
    Get #1, , BroachData(i)
Next i

Get #1, , BroachTeethDataCount
If BroachTeethDataCount > 0 Then
    ReDim BroachTeethData(BroachTeethDataCount) As BroachToothVars
    For i = 1 To BroachTeethDataCount
        Get #1, , BroachTeethData(i)
    Next i
End If

Close #1

BroachDataIndex = 0
RecordToTab1
RecordToTab2
RecordToTab3
RecordToTab4
RecordToTab5
RecordToTab6
RecordToTooth
End Sub

