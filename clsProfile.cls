VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sht_PF As Object
Private CAD As New clsACAD
Private f As New clsMyfunction
Private Math As New clsMath

Private collSpecificRow As New Collection

Private X0 As Double
Private Y0 As Double

Private Min As Double
Private Max As Double

Public Xscale As Integer
Public Yscale As Integer

Private sc As Integer
Private ec As Integer
Private txtheight As Double

Private Sub Class_Initialize()

Set sht_PF = Sheets("縱斷面")

With sht_PF

sc = 2
ec = .Cells(1, .Columns.Count).End(xlToLeft).Column

txtheight = 3

End With

Call CompareHeight

End Sub

Function test_getLoc(ByVal Xstep As Integer)

Dim coll As New Collection
Dim f As New clsMyfunction
Dim Math As New clsMath

With sht_PF

Lc = .Cells(1, .Columns.Count).End(xlToLeft).Column

firstloc = Math.TranLoc(.Cells(1, sc).Value)

For c = sc To ec Step Xstep

    myloc = Math.TranLoc(.Cells(1, c).Value)
    coll.Add myloc

Next

End With

Set test_getLoc = coll

End Function


Private Sub CompareHeight()

Dim collH As New Collection

With sht_PF

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 1 To lr

    heightname = .Cells(r, 1)
    
    If heightname Like "*高" Then
    
        For c = sc To ec
            If .Cells(r, c) <> 0 Then: collH.Add .Cells(r, c).Value
        Next
    
    End If

Next

Set collH_sort = f.BubbleSort_coll(collH)

Max = Int(collH_sort(collH_sort.Count)) + 1
Min = Int(collH_sort(1)) - 1

End With

End Sub

Sub setProperties(ByVal PF_point)

With sht_PF
    
    X0 = PF_point(0)
    Y0 = PF_point(1)

End With

End Sub

Sub getHeights()

With sht_PF

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 1 To lr

    heightname = .Cells(r, 1)

    If heightname Like "*高" Then

        Call getHeightByRow(r)

    End If

Next

End With

End Sub

Private Sub getHeightByRow(ByVal r)

Dim IsDrawStarted As Boolean
Dim collX As New Collection
Dim collY As New Collection

With sht_PF

HeightLayer = .Cells(r, 1)
firstloc = .Cells(1, sc)

For c = sc To ec

    nowH = .Cells(r, c)
    nowLoc = .Cells(1, c)

    If nowH <> 0 Then
    
        IsDrawStarted = True
        collX.Add X0 + (nowLoc - firstloc) * 1000 / Xscale
        collY.Add Y0 + (nowH - Min) * 1000 / Yscale
    
    ElseIf nowH = 0 And IsDrawStarted = True Then

        Set PF_Height = CAD.AddLWPolyLine(f.tranColls2OneArray(collX, collY))
        Call CAD.setLayer(PF_Height, "縱斷面-" & .Cells(r, 1))
        
        Set collX = Nothing: Set collY = Nothing
        
        IsDrawStarted = False

    End If

Next

If IsDrawStarted = True Then

        Set PF_Height = CAD.AddLWPolyLine(f.tranColls2OneArray(collX, collY))
        Call CAD.setLayer(PF_Height, "縱斷面-" & .Cells(r, 1))
        
    IsDrawStarted = False

End If

End With

End Sub


Sub DrawHeightBar() 'ByRef viewpoint As Variant)

Dim BaseHeightPoint(2) As Double
Dim spt(2) As Double, ept(2) As Double, txtpt(2) As Double

BaseHeightPoint(0) = X0 'viewpoint(0) - Interval
BaseHeightPoint(1) = Y0 'viewpoint(1) + (Interval + VHeight) * (TableRow)

BarStepHeight = 10 'mm

NeedStep = Int(((Max - Min) * 1000 / Yscale) / BarStepHeight)

With CAD 'controlobject

    For i = 1 To NeedStep
    
        spt(0) = BaseHeightPoint(0)
        spt(1) = BaseHeightPoint(1) + i * BarStepHeight
        
        ept(0) = spt(0) - 15
        ept(1) = spt(1)
        
        Set lineObj = .AddLine(spt, ept)
    
        Call CAD.setLayer(lineObj, "縱斷面-高程條")
    
        Set xlineObj = CAD.AddXLine(spt, ept)
        Call CAD.setLayer(xlineObj, "XLINE")
        'xlineObj.Layer = "XLINE"

        'lineObj.Layer = "縱斷面-高程條"
    
        txtpt(0) = BaseHeightPoint(0) - 20
        txtpt(1) = BaseHeightPoint(1) + i * BarStepHeight
    
        s = Min + i * Yscale / (1000 / BarStepHeight)
    
        Set txtobj = PlotTextCenter(s, txtpt, txtheight, 2)
        
        'txtobj.Layer = "縱斷面-高程條"
        Call CAD.setLayer(txtobj, "縱斷面-高程條")
        
        
    Next
    
    ept(0) = spt(0): ept(1) = spt(1)
    spt(0) = spt(0): spt(1) = BaseHeightPoint(1) + BarStepHeight
    
    Set lineObj = .AddLine(spt, ept)
    
    lineObj.Layer = "縱斷面-高程條"
    
    For i = 1 To NeedStep - 1
    
        spt(0) = BaseHeightPoint(0)
        spt(1) = BaseHeightPoint(1) + (i + 0.5) * BarStepHeight
        
        ept(0) = spt(0) - 6
        ept(1) = spt(1)
        
        Set lineObj = .AddLine(spt, ept)
        lineObj.Layer = "縱斷面-高程條"
    
        txtpt(0) = BaseHeightPoint(0) - 12
        txtpt(1) = BaseHeightPoint(1) + i * 10 + 5
    
        s = Min + (i + 0.5) * Yscale / (1000 / BarStepHeight)
    
        Set txtobj = PlotTextCenter(s, txtpt, txtheight, 2)
        txtobj.Layer = "縱斷面-高程條"
    
    Next

End With

End Sub

Sub DrawXBar(ByVal Xstep As Integer)

Dim spt(2) As Double
Dim ept(2) As Double
Dim txtpt(2) As Double

Set collLoc = test_getLoc(ByVal Xstep)

For Each it In collLoc

    spt(0) = X0 + (it - collLoc(1)) * 1000 / Xscale
    
    spt(1) = Y0 '- step / 2 * 1000
    ept(0) = spt(0)
    ept(1) = spt(1) - 3 '00 '* (PaperScale / 100)
    
    Set lineObj = CAD.AddLine(spt, ept)
    lineObj.Layer = "縱斷面-高程條"
    
    Set xlineObj = CAD.AddXLine(spt, ept)
    xlineObj.Layer = "XLINE"
    
    txtpt(0) = spt(0)
    txtpt(1) = ept(1) - 5 '00 '* (PaperScale / 100)
    
    Set txtobj = CAD.AddText(it, txtpt, 3, 2)
    txtobj.rotate CAD.tranPoint(txtpt), 3.14 / 2
    txtobj.Layer = "縱斷面-高程條"
    
Next

spt(0) = X0 + (collLoc(1) - collLoc(1)) * 1000 / Xscale
spt(1) = Y0
ept(0) = X0 + (collLoc(collLoc.Count) - collLoc(1)) * 1000 / Xscale
ept(1) = spt(1)

Set lineObj = CAD.AddLine(spt, ept)
lineObj.Layer = "縱斷面-高程條"

End Sub

'Sub getSpecificRow()
'
'Dim coll As New Collection
'
'With sht_PF
'
'    sr = 2
'    lr = .Cells(100, 1).End(xlUp).Row
'
'    For r = sr To lr
'
'        TargetName = .Cells(r, 1)
'
'        Select Case TargetName
'
'            Case "計畫高": rPlan = r
'            Case "坡降": rSlope = r
'            Case "設計起始點": rSE = r
'            Case "說明": rNote = r
'            Case "挖填方": rdeltaH = r
'
'        End Select
'
'    Next
'
'    If rPlan = 0 Then rPlan = InputBox("請輸入計畫高的列")
'    If rSlope = 0 Then rSlope = InputBox("請輸入坡度的列")
'    If rSE = 0 Then rSE = InputBox("請輸入設計起始點的列")
'    If rNote = 0 Then rNote = InputBox("請輸入說明的列")
'    If rdeltaH = 0 Then rdeltaH = InputBox("請輸入挖填方的列")
'
'    coll.Add rPlan, "rPlan"
'    coll.Add rSlope, "rSlope"
'    coll.Add rSE, "rSE"
'    coll.Add rNote, "rNote"
'    coll.Add rdeltaH, "rdeltaH"
'
'End With
'
'Set collSpecificRow = coll
'
'End Sub
