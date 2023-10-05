VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PI As Double

Function getAz(myX1, myY1, myX2, myY2) '取得方位角(以正北為0,順時+)

dx = myX2 - myX1
dy = myY2 - myY1
'PI = 4 * Atn(1)

If dx <> 0 Then thida = Atn(dy / dx) / PI * 180

If dx > 0 And dy > 0 Then '一
    fi = 90 - thida
ElseIf dx > 0 And dy < 0 Then '四
    fi = 90 - thida
ElseIf dx < 0 And dy < 0 Then '三
    fi = 270 - thida
ElseIf dx < 0 And dy > 0 Then '二
    fi = 270 - thida
ElseIf dx = 0 And dy > 0 Then '+Y
    fi = 0
ElseIf dx = 0 And dy < 0 Then '-Y
    fi = 180
ElseIf dx > 0 And dy = 0 Then '+X
    fi = 90
ElseIf dx < 0 And dy = 0 Then '-X
    fi = 270
End If

getAz = fi

End Function

Function getLength(ByVal spt, ByVal ept)

getLength = ((ept(0) - spt(0)) ^ 2 + (ept(1) - spt(1)) ^ 2) ^ 0.5

End Function

Function getLengthCO(ByVal X1, ByVal Y1, ByVal X2, ByVal Y2)

getLengthCO = ((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2) ^ 0.5

End Function

Function getCentroid(ByVal obj, Optional ByVal co As Double)

Dim Xc As New Collection
Dim Yc As New Collection
Dim Lc As New Collection
Dim scpt(2) As Double

'arr = obj.coordinates

If TypeName(obj) Like "*LWPolyline" Then
    arr = obj.coordinates
    coor = 2
ElseIf TypeName(obj) Like "*Polyline" Then
    arr = obj.coordinates
    coor = 3
ElseIf TypeName(obj) = "Double()" Then
    arr = obj
    coor = co
Else
    Debug.Print TypeName(obj)
    MsgBox "This function only support Polyline and LWPolyline", vbCritical: Exit Function
End If

For i = 0 To (UBound(arr) - coor) Step coor

    X1 = arr(i)
    Y1 = arr(i + 1)
    X2 = arr(i + coor)
    Y2 = arr(i + coor + 1)
    
    Xavg = (X1 + X2) / 2: Yavg = (Y1 + Y2) / 2
    L = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
    
    Xc.Add Xavg
    Yc.Add Yavg
    Lc.Add L

Next

For j = 1 To Xc.Count

    Xavg = Xc.Item(j)
    Yavg = Yc.Item(j)
    L = Lc.Item(j)
    
    Xw = Xw + Xavg * L
    Yw = Yw + Yavg * L
    LW = LW + L
    
Next

scpt(0) = Xw / LW: scpt(1) = Yw / LW

getCentroid = scpt

End Function

Function degcos(fi)

'PI = 4 * Atn(1)
thida = (90 - fi) * PI / 180
degcos = Cos(thida)

End Function

Function degsin(fi)

'PI = 4 * Atn(1)
thida = (90 - fi) * PI / 180
degsin = Sin(thida)

End Function

Function rad2deg(rad)

'PI = 4 * Atn(1)
rad2deg = rad / PI * 180

End Function

Function deg2rad(deg)

'PI = 4 * Atn(1)
deg2rad = deg / 180 * PI

End Function

Function TranLoc(ByVal Data As String) As Double

'樁號型態轉成可計算之樁號

tmp = Split(Data, "+")

If UBound(tmp) = -1 Or Data = "" Then Exit Function ' TranLoc = CDbl(Data): Exit Function

tloc = tmp(0) '千位數
dloc = tmp(1)

If dloc Like "*(*" Then

    tmp2 = Split(dloc, "(")

    tmp3 = Split(tmp2(0), ".")

    dloc = tmp3(0) + tmp3(1) / 10
    
End If

For i = 1 To Len(tloc)

    ch = mid(tloc, i, 1)
    If IsNumeric(ch) Then ref = ref & ch

Next

TranLoc = CDbl(ref) * 1000 + CDbl(dloc)
    
End Function

'Function collToArray(ByVal coll, Optional s As Double = 0)
'
'Dim tmp() As Variant
'
'Select Case s
'
'Case 0
'
'    ReDim tmp(0 To coll.Count - 1)
'
'    For i = 1 To coll.Count
'        tmp(i - 1) = coll(i)
'    Next
'
'Case 1
'
'    ReDim tmp(1 To coll.Count)
'
'    For i = 1 To coll.Count
'        tmp(i) = coll(i)
'    Next
'
'End Select
'
'collToArray = tmp
'
'End Function

Private Sub Class_Initialize()

PI = 4 * Atn(1)

End Sub

Function IsMiddle(ByVal X0 As Double, ByVal Y0 As Double, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double)

    midX = X1
    beforeX = X0
    afterX = X2
    midY = Y1
    beforeY = Y0
    afterY = Y2

    If (midX >= beforeX And midX <= afterX) Or (midX <= beforeX And midX >= afterX) Then
    
        XCheck = True
    
    Else
    
        XCheck = False
    
    End If
    
    If (midY >= beforeY And midY <= afterY) Or (midY <= beforeY And midY >= afterY) Then
    
        YCheck = True
    
    Else
    
        YCheck = False
    
    End If
    
    If (midY = beforeY And midY = afterY) And (midX = beforeX And midX = afterX) Then
    
        YCheck = True
        YCheck = True
    
    End If
    
    IsMiddle = False
    
    If XCheck = True And YCheck = True Then IsMiddle = True
    
End Function

Function interpolation(TargetLoc, sLoc, sh, eLoc, eH)

        If TargetLoc = sLoc Then
            interpolation = sh
        ElseIf TargetLoc = eLoc Then
            interpolation = eH
        ElseIf TargetLoc > sLoc And TargetLoc < eLoc Then
            interpolation = Round(((TargetLoc - sLoc) * eH + (eLoc - TargetLoc) * sh) / (eLoc - sLoc), 2)
        End If

End Function

Function Convert_Degree(Decimal_Deg) As Variant

    With Application
        'Set degree to Integer of Argument Passed
        Degrees = Int(Decimal_Deg)
        'Set minutes to 60 times the number to the right
        'of the decimal for the variable Decimal_Deg
        minutes = (Decimal_Deg - Degrees) * 60
        'Set seconds to 60 times the number to the right of the
        'decimal for the variable Minute
        seconds = Format(((minutes - Int(minutes)) * 60), "0")
        'Returns the Result of degree conversion
       '(for example, 10.46 = 10~ 27  ' 36")
        Convert_Degree = " " & Degrees & "° " & Int(minutes) & "' " _
            & seconds + Chr(34)
    End With
    
End Function

