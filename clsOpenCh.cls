VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenCh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sht As Worksheet
Private shtReport As Worksheet
Private activeRow As Integer

'Geometry

Private B As Double
Private m As Double
Private N As Double
Private H As Double
Private s As Double
Private f As Double
Private ls As Double
Private rs As Double

'Location

Private sLoc As String
Private eLoc As String

'Flow

Private DQ As Double

'Calc
Private manning As Double
Private d As Double
Private a As Double
Private p As Double
Private R_open As Double
Private n_open As Double
Private V As Double
Private q As Double


'水理因素表計算
'水理因素產生報告(PDF)

Private Sub Class_Initialize()

Set sht = Sheets("水理因素表")
Set shtReport = Sheets("水理(空白)")

End Sub

Sub getPropertiesByRow(ByVal r As Integer)

With sht

tmp = Split(.Cells(r, "A"), "~")
sLoc = tmp(0)
eLoc = tmp(1)
DQ = .Cells(r, "C")
manning = .Cells(r, "I")
s = .Cells(r, "M")
B = .Cells(r, "N")
m = .Cells(r, "O")
N = .Cells(r, "P")
f = .Cells(r, "Q")

End With

Call KeyInBasicParameter

End Sub

Private Sub KeyInBasicParameter()

With shtReport
    
    .Range("C5") = sLoc
    .Range("E5") = eLoc
    .Range("C7") = DQ
    .Range("C8") = manning
    .Range("C9") = s
    .Range("C10") = B
    .Range("G9") = f
    .Range("G7") = m
    .Range("G8") = N
    
End With

End Sub

Sub Calc_Report()

With shtReport
    
    For dh = 0.1 To 2 Step 0.01
        
        .Range("D15") = dh
        q_result = .Range("D26")
        q_check = .Range("D27")
        
        If q_result >= q_check Then Exit Sub
            
    Next
    
End With

End Sub

Sub ToPDF(ByVal FolderName As String, ByVal Filename As String, ByVal PrintSheetName As String) '給包商參考用的

wbpath = ThisWorkbook.Path

Sheets(PrintSheetName).Activate

    On Error Resume Next
    MkDir wbpath & "\" & FolderName
    On Error GoTo 0
    'ChDir wbpath
    
    On Error Resume Next
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    wbpath & "\" & FolderName & "\" & Filename, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
    
    If Err <> 0 Then
        MsgBox "你的excel版本過舊電腦無法另存為PDF!" & vbNewLine & "請點選列印紙張!": End
        Exit Sub
        
    End If
        
    On Error GoTo 0

End Sub


Sub Calc_ActiveRow(ByVal r As Integer)

With sht
    
    For dh = 0.1 To 2 Step 0.01
        
        .Cells(r, "L") = dh
        q_result = .Cells(r, "K")
        q_check = .Cells(r, "C")
        
        If q_result >= q_check Then Exit Sub
            
    Next
    
End With

End Sub

Sub getLocAndSlope() 'from Longitudinal Table

Dim coll As New Collection 'Loc
Dim coll2 As New Collection 'Slope
Dim s As String
Dim E As String

With Sheets("縱斷面繪圖")

    Lc = .Cells(2, 1).End(xlToRight).Column
    
    For c = 2 To Lc
    
        If .Cells(7, c) <> "" Then
        
            myloc = .Cells(2, c)
            slope = .Cells(7, c)
            
            coll.Add myloc
            coll2.Add slope
        
        End If
    
    Next

End With

With sht

r = 3

For i = 1 To coll.Count - 1

    s = coll(i)
    E = coll(i + 1)
    myslope = coll2(i)
    
    .Cells(r, "A") = s & "~" & E
    .Cells(r, "M") = myslope
    
    r = r + 1
    
Next

End With

End Sub

Sub Calc(ByVal myrow)

For d = 0.1 To H Step 0.01

    a = (B + B + m * d + N * d) * d / 2
    p = B + (1 + m ^ 2) ^ 0.5 * d + (1 + N ^ 2) ^ 0.5 * d
    R_open = a / p
    V = 1 / n_open * R_open ^ (2 / 3) * s ^ 0.5
    q = V * a
    
    Debug.Print q
    
    If q > DQ Then Call KeyInData(myrow): Exit For

Next

End Sub

Sub KeyInData(ByVal r As Double)

With Sheets("標準斷面")

    .Cells(r, "M") = Round(a, 3)
    .Cells(r, "N") = Round(p, 3)
    .Cells(r, "O") = Round(R_open, 3)
    .Cells(r, "P") = Round(R_open ^ (2 / 3), 3)
    .Cells(r, "Q") = Round(s ^ 0.5, 3)
    .Cells(r, "S") = Round(V, 3)
    .Cells(r, "T") = Round(q, 3)
    .Cells(r, "U") = Round(d, 2)

End With

End Sub
