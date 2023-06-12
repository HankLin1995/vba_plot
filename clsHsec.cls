VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHsec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private EGVertices
Private EG As Object
Private sht_EG As Object
Private CAD As New clsACAD
Private math As New clsMath
Private myFunc As New clsMyfunction

Private BaseX As Double
Private BaseY As Double
Private EG_sr As Double
Private EG_er As Double

Public PaperScale As Integer
Public SectionWidth As Integer
Public Is3D As Boolean

Private leftWidth As Double
Private rightWidth As Double
Private leftENVcode As String
Private rightENVcode As String

Private MidHeight As Double
Private step As Integer
Private loc As String

'TODO:
'1.繪製地界線
'2.繪製新設構造物

Private Sub Class_Initialize()

Set sht_EG = Sheets("橫斷面")
Is3D = False

End Sub

Sub setBaseHeightPT(ByVal BaseHeightPoint) 'ByVal bX As Double, ByVal bY As Double)

BaseX = BaseHeightPoint(0)
BaseY = BaseHeightPoint(1)

End Sub

Sub getPropertiesByRows(ByVal sr As Integer, ByVal er As Integer)

EG_sr = sr
EG_er = er

With sht_EG

    loc = .Cells(sr, 4)
    
    leftENVcode = .Cells(sr, 3)
    rightENVcode = .Cells(er, 3)
    
    leftWidth = Int(.Cells(sr, 1)) - 2
    rightWidth = Int(.Cells(er, 1)) + 2

End With

'BaseX = BaseX - leftWidth * 1000 '重新調整邊界X

If Is3D = False Then Call getMidHeight(sr, er)
Call getStep(sr, er)
Call getEGVertices(sr, er)

'=======校正形心位置========

If Is3D = False Then

    a = math.getCentroid(EGVertices, 2)
    distanceY = Round((a(1) - BaseY) / 1000)
    
    If distanceY <> 0 Then
    
        MidHeight = MidHeight + distanceY
        Call getEGVertices(sr, er)
    
    End If

End If

'===========================

End Sub

Private Sub getEGVertices(ByVal sr As Integer, ByVal er As Double)

'需要先得到MidHeight

'判斷條件
'1.為0時是否要顯示
'2.特定地界線是否要顯示

Dim collX As New Collection
Dim collY As New Collection

With sht_EG
    
    For r = sr To er
    
        dx = .Cells(r, 1)
        dy = .Cells(r, 2)
        
        If dy <> 0 Then
        
            collX.Add BaseX + dx * 1000
            collY.Add BaseY + (dy - MidHeight) * 1000

        End If
        
    Next

End With

'=======ExtendEG============

'Call collX.Add(collX(1) - 1000, , 1)
'Call collY.Add(collY(1), , 1)
'
'Call collX.Add(collX(collX.Count) + 1000)
'Call collY.Add(collY(collY.Count))

EGVertices = myFunc.tranColls2OneArray(collX, collY)

End Sub

Private Sub getMidHeight(ByVal sr As Integer, ByVal er As Integer)

'繪製Y-bar的時候先將橫斷面中心點移動至各樁號設定位置

With sht_EG

For r = sr To er

    dataH = Val(.Cells(r, 2))

    If dataH <> 0 Then
        
        sumH = sumH + dataH
        cnt = cnt + 1
    
    End If
    
Next

MidHeight = Int(sumH / cnt)

End With

End Sub

Function getStep(ByVal sr As Integer, ByVal er As Integer)

Dim coll As New Collection
Dim myFunc As New clsMyfunction

With sht_EG

For r = sr To er

    dataH = Val(.Cells(r, 2))
    
    If dataH <> 0 Then coll.Add dataH
    
Next

Set o = myFunc.BubbleSort_coll(coll)

step = Int(o(o.Count) - o(1)) + 2

getStep = step

End With

End Function

Sub plotHsec()

Set EG = CAD.AddLWPolyLine(EGVertices)

EG.Layer = "GL" '"橫斷面-原地面線"

End Sub

Sub plotUchannel()

Dim cpt(2) As Double

With sht_EG

    For r = EG_sr + 1 To EG_er

        If .Cells(r, 4) <> "" Then
        
            dx = .Cells(r, 1)
            args = Split(.Cells(r, 4), ",")
            planHeight = args(0)
            planType = args(1)
        
            Dim Uobj As New clsUChannel
            
            Uobj.loc = loc
            cpt(0) = BaseX + dx * 1000
            cpt(1) = BaseY + (planHeight - MidHeight) * 1000
            
            Uobj.CenterPoint = cpt
            Uobj.ReadDataFromXLS (planType)
            Uobj.DrawMainStructure
            Call Uobj.plotDigLine(EG, EGVertices)
        
        End If
        
    Next

End With

End Sub

Sub plotUchannel_Block()

Dim cpt(2) As Double

With sht_EG

    For r = EG_sr + 1 To EG_er

        If .Cells(r, 4) <> "" Then
        
            dx = .Cells(r, 1)
            args = Split(.Cells(r, 4), ",")
            planHeight = args(0)
            planType = args(1)
        
            Dim Uobj As New clsUChannel
            
            Uobj.loc = loc
            cpt(0) = BaseX + dx * 1000
            cpt(1) = BaseY + (planHeight - MidHeight) * 1000
            
            Call CAD.InsertBlock(cpt, "標準斷面", 1)
            
            'Uobj.CenterPoint = cpt
            'Uobj.ReadDataFromXLS (planType)
            'Uobj.DrawMainStructure
            'Call Uobj.plotDigLine(EG, EGVertices)
        
        End If
        
    Next

End With

End Sub

Sub plotOther(ByVal collENV)

Dim LeftGroundPt(2) As Double
Dim RightGroundPt(2) As Double

LeftGroundPt(0) = EGVertices(0)
LeftGroundPt(1) = EGVertices(1)

Set o = CAD.InsertBlock(LeftGroundPt, "LeftGL", PaperScale, Array(collENV(leftENVcode)))
o.Layer = "GL" ' "橫斷面-原地面線"

RightGroundPt(0) = EGVertices(UBound(EGVertices) - 1)
RightGroundPt(1) = EGVertices(UBound(EGVertices))

Set o = CAD.InsertBlock(RightGroundPt, "RightGL", PaperScale, Array((collENV(rightENVcode))))
o.Layer = "GL" ' "橫斷面-原地面線"

End Sub

Sub plotTitle()

Dim hcpt(2) As Double
Dim coll As New Collection

hcpt(0) = BaseX + CDbl(SectionWidth) * 1000
hcpt(1) = BaseY + step / 2 * 1000 + 1500 * (PaperScale / 100)

coll.Add loc

Set o = CAD.InsertBlock(hcpt, "H_TITLE", PaperScale, coll)

o.Layer = "TITLE"

End Sub

Sub DrawHeightBar() 'plot the y-height-bar，橫斷面高程表可以再做精緻些

'CAD.acaddoc.ActiveLayer = CAD.acaddoc.Layers("橫斷面-高程條")
CAD.acaddoc.ActiveLayer = CAD.acaddoc.Layers("BAR")

Dim spt(2) As Double
Dim ept(2) As Double
Dim txtpt(2) As Double

spt(0) = BaseX + leftWidth * 1000

For i = -step / 2 To step / 2 'the range of up and down is step/2
    
    spt(1) = BaseY + i * 1000
    ept(1) = spt(1)
    ept(0) = spt(0) - 600 * (PaperScale / 100)
    
    Set lineObj = CAD.AddLine(spt, ept)
    
    txtpt(0) = ept(0) - 200 * (PaperScale / 100)
    txtpt(1) = ept(1)

    Set txtobj = CAD.AddText(MidHeight + i, txtpt, PaperScale * 2, 3)

Next

For i = -step / 2 To step / 2 - 1

    spt(1) = BaseY + (i + 0.5) * 1000
    ept(1) = spt(1)
    ept(0) = spt(0) - 300 * (PaperScale / 100)
    
    Set lineObj = CAD.AddLine(spt, ept)
    
    txtpt(0) = ept(0) - 200 * (PaperScale / 100)
    txtpt(1) = ept(1)
    
    Set txtobj = CAD.AddText(MidHeight + i + 0.5, txtpt, PaperScale * 2, 3)
    
Next

spt(0) = BaseX + leftWidth * 1000
spt(1) = BaseY - step / 2 * 1000
ept(0) = BaseX + leftWidth * 1000
ept(1) = BaseY + step / 2 * 1000

Set lineObj = CAD.AddLine(spt, ept)

CAD.acaddoc.ActiveLayer = CAD.acaddoc.Layers("0")

Call DrawXBar
Call ExportBorderToCL

End Sub

Private Sub DrawXBar() 'plot the y-height-bar，橫斷面高程表可以再做精緻些

'CAD.acaddoc.ActiveLayer = CAD.acaddoc.Layers("橫斷面-高程條")
CAD.acaddoc.ActiveLayer = CAD.acaddoc.Layers("BAR")

Dim spt(2) As Double
Dim ept(2) As Double
Dim txtpt(2) As Double

'spt(0) = BaseX + EGWidth * 1000

For i = leftWidth To rightWidth

    spt(0) = BaseX + (i) * 1000
    spt(1) = BaseY - step / 2 * 1000
    ept(0) = spt(0)
    ept(1) = spt(1) - 300 * (PaperScale / 100)
    
    Set lineObj = CAD.AddLine(spt, ept)
    
    txtpt(0) = spt(0)
    txtpt(1) = ept(1) - 200 * (PaperScale / 100)
    
    Set txtobj = CAD.AddText(i - SectionWidth, txtpt, PaperScale * 2, 2)
    
Next

For i = leftWidth To rightWidth - 1

    spt(0) = BaseX + (i + 0.5) * 1000
    spt(1) = BaseY - step / 2 * 1000
    ept(0) = spt(0)
    ept(1) = spt(1) - 150 * (PaperScale / 100)

    Set lineObj = CAD.AddLine(spt, ept)

    'txtpt(0) = spt(0)
    'txtpt(1) = ept(1) - 200 * (PaperScale / 100)

    'Set txtobj = CAD.AddText(i + 0.5, txtpt, PaperScale * 2, 2)
Next

spt(0) = BaseX + leftWidth * 1000
spt(1) = BaseY - step / 2 * 1000
ept(0) = BaseX + (rightWidth) * 1000
ept(1) = spt(1)

Set lineObj = CAD.AddLine(spt, ept)

CAD.acaddoc.ActiveLayer = CAD.acaddoc.Layers("0")

End Sub

Private Sub ExportBorderToCL()

MinX = BaseX + leftWidth * 1000
MinY = BaseY - step / 2 * 1000
MaxX = BaseX + (rightWidth) * 1000
MaxY = BaseY + step / 2 * 1000

With Sheets("中心線")
    
    Set rng = .Columns(1).Find(loc)
    
    If Not rng Is Nothing Then
    
    r = rng.Row
    .Cells(r, 3) = Join(Array(MinX, MinY, MaxX, MaxY), ",")

    End If

End With

End Sub
