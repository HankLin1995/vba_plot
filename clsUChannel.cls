VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public CenterPoint As Variant
Public loc As String

Private BasicParameterArr As Variant
Private ControlCalculation As Variant

Private BW As Double
Private BT As Double
Private LH As Double
Private LT As Double
Private RH As Double
Private RT As Double
Private OW As Double
Private BBT As Double
Private UF As Double
Private DF As Double

Private MainCONC As Object
Private buttomCONC As Object

Private CAD As New clsACAD
Private myFunc As New clsMyfunction
Private sht As Object
Private shtChannelType As Object

Private Sub Class_Initialize()

Set sht = Sheets("橫斷面")
Set shtChannelType = Sheets("渠道型式表")
Set ACAD = CAD.acadDoc
Set mo = ACAD.ModelSpace

End Sub

Sub ReadDataFromXLS(ByVal TargetType As String)

Multiple = 1000
Ratio = 1
c = 8

With shtChannelType

    For r = 3 To .Cells(Rows.Count, c).End(xlUp).Row
    
        If .Cells(r, c - 1) = TargetType Then Exit For
    Next

    BW = .Cells(r, c + 0) * Multiple / Ratio
    BT = .Cells(r, c + 1) * Multiple / Ratio
    LH = .Cells(r, c + 2) * Multiple / Ratio
    LT = .Cells(r, c + 3) * Multiple / Ratio
    RH = .Cells(r, c + 4) * Multiple / Ratio
    RT = .Cells(r, c + 5) * Multiple / Ratio
    OW = .Cells(r, c + 6) * Multiple / Ratio
    BBT = .Cells(r, c + 7) * Multiple / Ratio
    UF = Val(.Cells(r, c + 8)) * Multiple / Ratio
    DF = Val(.Cells(r, c + 9)) * Multiple / Ratio
    
    DimParameterArr = Array(RH, RT, BW, LT, LH, BT, BBT)

    If DF = 0 Then
    
    BasicParameterArr = Array(BW / 2, LH, LT, LH, OW, BT, OW + LT + BW + RT + OW, BT, OW, RH, RT, RH, BW / 2)
    ControlCalculation = Array("-X", "+Y", "-X", "-Y", "-X", "-Y", "+X", "+Y", "-X", "+Y", "-X", "-Y", "-X")
    
    Else
    
    BasicParameterArr = Array(BW / 2 - DF, DF, LH - DF, LT, LH, OW, BT, OW + LT + BW + RT + OW, BT, OW, RH, RT, RH - DF, DF, BW / 2 - DF)
    ControlCalculation = Array("-X", "LU", "+Y", "-X", "-Y", "-X", "-Y", "+X", "+Y", "-X", "+Y", "-X", "-Y", "LD", "-X")
    
    End If
    
End With

End Sub

Sub DrawMainStructure()

Dim vertices() As Double
Dim spt(2) As Double
Dim Ldpt(2) As Double, rupt(2) As Double
ReDim vertices(3 * (UBound(BasicParameterArr) + 2) - 1) '頭尾各加一

spt(0) = CenterPoint(0): spt(1) = CenterPoint(1)

vertices(0) = spt(0)
vertices(1) = spt(1)

For i = 0 To UBound(BasicParameterArr)
    
    Calc = ControlCalculation(i)
    Shape = BasicParameterArr(i)
    
    Select Case Calc
    
    Case "+X": spt(0) = spt(0) + Shape
    Case "-X": spt(0) = spt(0) - Shape
    Case "+Y": spt(1) = spt(1) + Shape
    Case "-Y": spt(1) = spt(1) - Shape
    
    Case "LU"
    
        spt(0) = spt(0) - Shape
        spt(1) = spt(1) + Shape
    
    Case "LD"
    
        spt(0) = spt(0) - Shape
        spt(1) = spt(1) - Shape
    
    End Select

    vertices(3 + i * 3) = spt(0)
    vertices(4 + i * 3) = spt(1)
    
Next

Set MainCONC = CAD.AddPolyLine(vertices)

MainCONC.Layer = "CON" ' "橫斷面-結構線"

Ldpt(0) = CenterPoint(0) - BW / 2 - LT - OW
Ldpt(1) = CenterPoint(1) - BT - BBT
rupt(0) = CenterPoint(0) + BW / 2 + RT + OW
rupt(1) = CenterPoint(1) - BT

Set buttomCONC = CAD.PlotRec(Ldpt, rupt)

buttomCONC.Layer = "CON" ' "橫斷面-結構線"

Ldpt(0) = CenterPoint(0)
Ldpt(1) = CenterPoint(1) + (RH + LH) / 2 + 4 * OW
rupt(0) = CenterPoint(0)
rupt(1) = CenterPoint(1) - 3 * BT

Set lineObj = CAD.AddLine(Ldpt, rupt)

lineObj.Layer = "CL" ' "橫斷面-中心線"

End Sub

Sub plotDigLine(ByVal EG As Object, ByVal EGVertices)

Dim IsIntersect As Boolean

Call CAD.GetBoundingBox(buttomCONC, MinX, MinY, MaxX, MaxY)

Dim vertices(4 * 2 - 1) As Double

vertices(0) = MinX - 300 - 30
vertices(1) = MinY + 100
vertices(2) = MinX - 300
vertices(3) = MinY
vertices(4) = MaxX + 300
vertices(5) = MinY
vertices(6) = MaxX + 300 + 30
vertices(7) = MinY + 100

Set DL = CAD.AddLWPolyLine(vertices)
DL.Layer = "CAL" ' "橫斷面-挖方線"

retpt = myFunc.SortPTArray(CAD.IntersectWith(DL, EG, 1, IsIntersect))

If IsIntersect = False Then Exit Sub

vertices(0) = retpt(0, 0)
vertices(1) = retpt(0, 1)

vertices(6) = retpt(UBound(retpt, 1), 0) '- 3 + 1)
vertices(7) = retpt(UBound(retpt, 1), 1) '- 3 + 2)

'========以邊界為基準================

'vertices(0) = EGVertices(0)
'vertices(1) = EGVertices(1)
'vertices(6) = EGVertices(UBound(EGVertices) - 1)
'vertices(7) = EGVertices(UBound(EGVertices))

'====================================

DL.Delete

Set DL2 = CAD.AddLWPolyLine(vertices)
DL2.Layer = "CAL" '"橫斷面-挖方線"

'============匯入挖方面積==================

Dim stoneObj As New clsStone

DLArea = stoneObj.getDLArea(EG, DL2)

With Sheets("中心線")
    
    Set rng = .Columns(1).Find(loc)
    
    If Not rng Is Nothing Then
    
    r = rng.Row
    .Cells(r, 5) = DLArea
    
    End If

End With

End Sub

