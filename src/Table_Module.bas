Attribute VB_Name = "Table_Module"
Option Explicit

'查公差例程
'Dim U#, L#
'If IsNumeric(TextD.Text) Then
'    GetLimitFromTable Val(TextD.Text), ComboDToleranceZone.Text, U, L
'    TextDMax.Text = FixLimit(U / 1000)
'    TextDMin.Text = FixLimit(L / 1000)
'End If


' ***********************************************************************
'   目的：      获得公差等级
'
'   输入：      Tolerance：输入的公差字符串
'
'   例子：      输入"H7"返回7
' ***********************************************************************
Public Function GetToleranceGrade(Tolerance As String) As Integer
Dim i As Integer
For i = 1 To Len(Tolerance)
    If IsNumeric(Mid(Tolerance, i, 1)) = True Then
        GetToleranceGrade = Val(Mid(Tolerance, i, Len(Tolerance) - i + 1))
        Exit For
    End If
Next i
End Function

' ***********************************************************************
'   目的：      获得公差代号
'
'   输入：      Tolerance：输入的公差字符串
'
'   例子：      输入"H7"返回"H"
' ***********************************************************************
Public Function GetToleranceCode(Tolerance As String) As String
Dim i As Integer
For i = 1 To Len(Tolerance)
    If IsNumeric(Mid(Tolerance, i, 1)) = True Then
        GetToleranceCode = Mid(Tolerance, 1, i - 1)
        Exit For
    End If
Next i
End Function

' ***********************************************************************
'   目的：      修正公差：自动在数字前加正负号及0
'
'   输入：      Num：极限偏差
'
'   例子：
' ***********************************************************************
Public Function FixLimit(ByVal Num As Double) As String
FixLimit = Format(Num, "+0.000;-0.000;0")
'Select Case Num
'Case Is < -1
'    FixLimit = "-" & Num
'Case Is < 0
'    FixLimit = "-0" & Mid(Str(Num), 2, Len(Str(Num)) - 1)
'Case Is = 0
'    FixLimit = Num
'Case Is < 1
'    FixLimit = "+0" & Num
'Case Is >= 1
'    FixLimit = "+" & Num
'End Select
End Function
' ***********************************************************************
'   目的：      由公差带及孔轴径得出上下极限（未完）
'
'   输入：      D：直径
'               ToleranceZone：公差带
'               UpperLimitDeviation：引用传递，上极限偏差
'               LowerLimitDeviation：引用传递，下极限偏差
'
'   例子：
' ***********************************************************************
Public Sub GetLimitFromTable(D As Double, ToleranceZone As String, ByRef UpperLimitDeviation, LowerLimitDeviation As Double)
Dim ToleranceGrade As Integer
Dim ToleranceCode As String
Dim IsStandard As Boolean
Dim IsAxis As Boolean '是轴
Dim IsES As Boolean '上偏差
Dim IsEI As Boolean '下偏差
Dim TableName As String

Dim Connection1 As New ADODB.Connection
Dim Recordset1 As New ADODB.Recordset
Dim DatabaseName As String
DatabaseName = App.Path & "\拉刀设计数据库.mdb"
Connection1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseName

ToleranceCode = GetToleranceCode(ToleranceZone) '公差代号
ToleranceGrade = GetToleranceGrade(ToleranceZone) '公差等级
IsAxis = (Asc(ToleranceCode) >= Asc("a")) And (Asc(ToleranceCode) <= Asc("z")) '判断轴孔
IsES = (Asc(ToleranceCode) >= Asc("a")) And (Asc(ToleranceCode) <= Asc("h")) '上偏差
IsEI = (Asc(ToleranceCode) >= Asc("m")) And (Asc(ToleranceCode) <= Asc("z")) '下偏差
IsEI = IsEI Or (Asc(ToleranceCode) >= Asc("A")) And (Asc(ToleranceCode) <= Asc("H")) '下偏差
'轴：js,j,k特殊
If IsAxis Then 'a~z
    TableName = "轴的基本偏差"
Else
    TableName = "孔的基本偏差"
End If

If IsES Then '查表为上偏差 a~h
    Recordset1.Open "SELECT " & ToleranceCode & " FROM " & TableName & " WHERE " & D & ">大于 AND " & D & "<=至", Connection1, 1, 1 '根据D值及代号得到上下偏差
    UpperLimitDeviation = Recordset1.Fields(ToleranceCode)
    Recordset1.Close
    
    Recordset1.Open "SELECT IT" & ToleranceGrade & " FROM 标准公差数值表 WHERE " & D & ">大于 AND " & D & "<=至", Connection1, 1, 1  '根据D值及IT*得到数值
    LowerLimitDeviation = UpperLimitDeviation - Val(Recordset1.Fields("IT" & ToleranceGrade))
    Recordset1.Close
    
    If ToleranceGrade >= 12 Then '大于12级
        UpperLimitDeviation = UpperLimitDeviation * 1000
        LowerLimitDeviation = LowerLimitDeviation * 1000
    End If
End If

If IsEI Then '查表为下偏差 m~z,A~H
    Recordset1.Open "SELECT " & ToleranceCode & " FROM " & TableName & " WHERE " & D & ">大于 AND " & D & "<=至", Connection1, 1, 1 '根据D值及代号得到上下偏差
     LowerLimitDeviation = Recordset1.Fields(ToleranceCode)
    Recordset1.Close
    
    Recordset1.Open "SELECT IT" & ToleranceGrade & " FROM 标准公差数值表 WHERE " & D & ">大于 AND " & D & "<=至", Connection1, 1, 1  '根据D值及IT*得到数值
    UpperLimitDeviation = UpperLimitDeviation + Val(Recordset1.Fields("IT" & ToleranceGrade))
    Recordset1.Close
    
    If ToleranceGrade >= 12 Then '大于12级
        UpperLimitDeviation = UpperLimitDeviation * 1000
        LowerLimitDeviation = LowerLimitDeviation * 1000
    End If
End If

End Sub

' ***********************************************************************
'   目的：      判断表是否存在（未用）
'
'   输入：
'
'   例子：
' ***********************************************************************
Public Function TableExist(mdbName As String, TableName As String) As Boolean
Dim TableExit As Boolean
Dim cn As New ADODB.Connection, Rs As New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbName & ";Persist Security Info=False"
On Error Resume Next
Rs.Open "Select * From " & TableName, cn
If Err.Number = 0 Then
    TableExit = True
End If
Err.Clear
Rs.Close
cn.Close
Set Rs = Nothing
Set cn = Nothing
End Function

