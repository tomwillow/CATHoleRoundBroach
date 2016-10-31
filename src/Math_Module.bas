Attribute VB_Name = "Math_Module"
Option Explicit

'本页所有三角函数，及Math对象下三角函数各参数均为弧度

Public Function PI() As Double
PI = 4 * Atn(1#)
End Function

'Secant
Public Function Sec(ByVal Number As Double) As Double
    Sec = 1 / Cos(Number)
End Function

'Cosecant
Public Function Csc(ByVal Number As Double) As Double
    Csc = 1 / Sin(Number)
End Function

'Cotangent
Public Function Ctn(ByVal Number As Double) As Double
    Ctn = 1 / Tan(Number)
End Function

'Inverse Sine
Public Function ASin(ByVal Number As Double) As Double
    ASin = Atn(Number / Sqr(-Number * Number + 1))
End Function

'Inverse Cosine
Public Function ACos(ByVal Number As Double) As Double
    ACos = Atn(-Number / Sqr(-Number * Number + 1)) + 2 * Atn(1)
End Function

'Inverse Secant
Public Function ASec(ByVal Number As Double) As Double
    ASec = Atn(Number / Sqr(Number * Number - 1)) + Sgn((Number) - 1) * (2 * Atn(1))
End Function

'Inverse Cosecant
Public Function ACsc(ByVal Number As Double) As Double
    ACsc = Atn(Number / Sqr(Number * Number - 1)) + (Sgn(Number) - 1) * (2 * Atn(1))
End Function

'Inverse Cotangent
Public Function ACtn(ByVal Number As Double) As Double
    ACtn = Atn(Number) + 2 * Atn(1)
End Function

'Hyperbolic Sine
Public Function SinH(ByVal Number As Double) As Double
    SinH = (Exp(Number) - Exp(-Number)) / 2
End Function

'Hyperbolic Cosine
Public Function CosH(ByVal Number As Double) As Double
    CosH = (Exp(Number) + Exp(-Number)) / 2
End Function

'Hyperbolic Tangent
Public Function TanH(ByVal Number As Double) As Double
    TanH = (Exp(Number) - Exp(-Number)) / (Exp(Number) + Exp(-Number))
End Function

'Hyperbolic Secant
Public Function SecH(ByVal Number As Double) As Double
    SecH = 2 / (Exp(Number) + Exp(-Number))
End Function

'Hyperbolic Cosecant
Public Function CscH(ByVal Number As Double) As Double
    CscH = 2 / (Exp(Number) - Exp(-Number))
End Function

'Hyperbolic Cotangent
Public Function CtnH(ByVal Number As Double) As Double
    CtnH = (Exp(Number) + Exp(-Number)) / (Exp(Number) - Exp(-Number))
End Function

'Inverse Hyperbolic Sine
Public Function ASinH(ByVal Number As Double) As Double
    ASinH = Log(Number + Sqr(Number * Number + 1))
End Function

'Inverse Hyperbolic Cosine
Public Function ACosH(ByVal Number As Double) As Double
    ACosH = Log(Number + Sqr(Number * Number - 1))
End Function

'Inverse Hyperbolic Tangent
Public Function ATanH(ByVal Number As Double) As Double
    ATanH = Log((1 + Number) / (1 - Number)) / 2
End Function

'Inverse Hyperbolic Secant
Public Function ASecH(ByVal Number As Double) As Double
    ASecH = Log((Sqr(-Number * Number + 1) + 1) / Number)
End Function

'Inverse Hyperbolic Cosecant
Public Function ACscH(ByVal Number As Double) As Double
    ACscH = Log((Sgn(Number) * Sqr(Number * Number + 1) + 1) / Number)
End Function

'Inverse Hyperbolic Cotangent
Public Function ACtnH(ByVal Number As Double) As Double
    ACtnH = Log((Number + 1) / (Number - 1)) / 2
End Function

'Logarithm to base N
Public Function LogN(ByVal Base As Double, ByVal Number As Double)
    LogN = Log(Number) / Log(Base)
End Function

Public Function DegToRad(ByVal Deg As Double) As Double
DegToRad = Deg / 180 * PI
End Function

Public Function RadToDeg(ByVal Rad As Double) As Double
RadToDeg = Rad * 180 / PI
End Function

Public Function AriSquSum(ByVal a1#, ByVal an#, Optional ByVal D# = 1) As Double '等差数列求和
Dim n#
n = (an - a1) / D + 1
AriSquSum = n * (a1 + an) / 2
End Function

' ***********************************************************************
'   目的：      字符串转浮点数，若非数字弹出提示
'
'   输入：
' ***********************************************************************
Public Function ValF(S, Name As String, Optional DefaultNum As Double) As Double
If IsNumeric(S) Then
    ValF = Val(S)
Else
    MsgBox "请设定" & Name & "。"
    Form1.SendMsgStr "请设定" & Name & "。"
    ValF = DefaultNum
End If
End Function

' ***********************************************************************
'   目的：      上进至5
'
'   输入：
'
'   例子：      26.0125→26.015
' ***********************************************************************
Public Function UpTo5(Num As Double) As Double '
Num = Math.Round(Num, 3)
Select Case (Num * 100 - Int(Num * 100)) * 10
Case Is <= 2
    UpTo5 = Math.Round(Num, 2)
Case Is <= 5
    UpTo5 = Int(Num * 100) / 100 + 0.005
Case Is <= 9
    UpTo5 = Math.Round(Num, 2)
End Select
End Function

Public Function Fix0(ByVal S As String) As String '修正0：大于-1负数负号前加0，小于1正数小数点前加0
Fix0 = Format(S, "0.#######;-0.#######;0") '最大7位小数够了吧
'Select Case Val(S)
'Case Is < -1:
'    Fix0 = S
'Case Is < 0: '大于-1负数负号前加0
'    Fix0 = Mid(S, 1, 1) & "0" & Mid(S, 2)
'Case Is = 0:
'    Fix0 = S
'Case Is < 1: '小于1正数小数点前加0
'    Fix0 = "0" & S
'Case Else
'    Fix0 = S
'End Select
End Function

Public Function GetAngleFromPoint(ByVal iCenterX As Double, ByVal iCenterY As Double, ByVal X As Double, ByVal Y As Double) As Double
If ((X - iCenterX) > 0) And ((Y - iCenterY) >= 0) Then '第1象限[0,PI/2)
    GetAngleFromPoint = Atn((Y - iCenterY) / (X - iCenterX))
End If
If ((X - iCenterX) <= 0) And ((Y - iCenterY) > 0) Then '第2象限[PI/2,PI)
    GetAngleFromPoint = PI / 2 + Atn((iCenterX - X) / (Y - iCenterY))
End If
If ((X - iCenterX) < 0) And ((Y - iCenterY) <= 0) Then '第3象限[PI,PI*3/2)
    GetAngleFromPoint = PI + Atn((iCenterY - Y) / (iCenterX - X))
End If
If ((X - iCenterX) >= 0) And ((Y - iCenterY) < 0) Then '第4象限[PI*3/2,2*PI)
    GetAngleFromPoint = 3 / 2 * PI + Atn((X - iCenterX) / (iCenterY - Y))
End If
End Function
