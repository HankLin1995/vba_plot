VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelToCAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MyACAD As New clsACAD

Const TXT_COE = 0.6 '0.8
Const ROW_COE = 0.335
Const COL_COE = 2.1

Public lupt As Variant
Private rng_start As Range
Private rng_end As Range
Private rng_first As Range

Sub ExportToCAD()

For Each rng In Selection

    If c = 0 Then Set rng_first = rng

    If rng.MergeCells Then
    
        Set ma = rng.MergeArea
    
        ma.Interior.ColorIndex = 41
        MergeTmp = Split(ma.Address, ":")
        
        Set rng_start = Range(MergeTmp(0)): Set rng_end = Range(MergeTmp(1))
        
        If rng.Address = rng_start.Address Then
        
            Call DrawTable(rng_start, rng_end)
        
        End If
        
        If rng.Address = rng_end.Address Then rng_start.Interior.ColorIndex = 0
        
    Else
    
        Call DrawTable(rng, rng)
        'Call DrawBorder(rng)
    
    End If

    c = c + 1

Next

End Sub

Sub DrawTable(ByVal rng_start As Range, ByVal rng_end As Range)

Dim txtpt(2) As Double
Dim Ldpt(2) As Double
Dim rupt(2) As Double

txt_Height = rng_start.Font.size
txt_String = rng_start.Value

CW = GetCW(rng_start, rng_end)
RH = GetRH(rng_start, rng_end)

txtpt(0) = rng_start.LEFT + CW / 2 + lupt(0)
txtpt(1) = -rng_start.Top - RH / 2 + lupt(1)

Set txtobj = MyACAD.AddMixText(txt_String, txtpt, txt_Height * TXT_COE, 2)

If rng_start = rng_end Then

    Call DrawBorder(rng_start)

Else

    Ldpt(0) = rng_start.LEFT + lupt(0)
    Ldpt(1) = -rng_start.Top + lupt(1)
    rupt(0) = rng_start.LEFT + CW + lupt(0)
    rupt(1) = -rng_start.Top - RH + lupt(1)
    
    Set recobj = MyACAD.PlotRec(Ldpt, rupt)

End If

End Sub

Sub DrawBorder(ByVal rng As Range)

Dim vertices(2 * 3 - 1) As Double

CW = GetCW(rng, rng)
RH = GetRH(rng, rng)

If rng.Borders(xlEdgeLeft).LineStyle <> -4142 And rng.LEFT = rng_first.LEFT Then

    vertices(0) = rng.LEFT + lupt(0): vertices(1) = -rng.Top + lupt(1)
    vertices(3) = vertices(0): vertices(4) = vertices(1) - RH
    
    Set plineobj = MyACAD.AddPolyLine(vertices)
    
End If

If rng.Borders(xlEdgeTop).LineStyle <> -4142 And rng.Top = rng_first.Top Then

    vertices(0) = rng.LEFT + lupt(0): vertices(1) = -rng.Top + lupt(1)
    vertices(3) = vertices(0) + CW: vertices(4) = vertices(1)
    
    Set plineobj = MyACAD.AddPolyLine(vertices)
    
End If

If rng.Borders(xlEdgeRight).LineStyle <> -4142 Then

    vertices(0) = rng.LEFT + lupt(0) + CW: vertices(1) = -rng.Top + lupt(1)
    vertices(3) = vertices(0): vertices(4) = vertices(1) - RH
    
    Set plineobj = MyACAD.AddPolyLine(vertices)
    
End If

If rng.Borders(xlEdgeBottom).LineStyle <> -4142 Then

    vertices(0) = rng.LEFT + lupt(0): vertices(1) = -rng.Top + lupt(1) - RH
    vertices(3) = vertices(0) + CW: vertices(4) = vertices(1)
    
    Set plineobj = MyACAD.AddPolyLine(vertices)
    
End If

End Sub

Function GetCW(ByVal rng1 As Range, ByVal rng2 As Range)

GetCW = rng2.Offset(0, 1).LEFT - rng1.LEFT

End Function

Function GetRH(ByVal rng1 As Range, ByVal rng2 As Range)

GetRH = rng2.Offset(1, 0).Top - rng1.Top

End Function

Private Sub Class_Initialize()

lupt = MyACAD.GetPoint("請點選表格的左上角點")

End Sub
