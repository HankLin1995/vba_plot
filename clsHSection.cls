VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private BaseHeightPoint As Variant
Private XYTable As Variant
Private collLineCD As New Collection
Private MyACAD As New clsACAD
Private collTargetHeight As New Collection
Private collFinalHeight As New Collection
Private collLoc As New Collection
Private collPlanDiff As New Collection
Private colltmpSecLoc As New Collection
Private collPtTable As New Collection

Private Xscale As Double
Private Yscale As Double
Public PaperScale As Double
Public step As Integer '樁號說明距離
Public SectionWidth As Double
Public MoveY As Double '各個樁號距離
Public IsMidHeightFixed As Boolean
Public MidHeight As Integer
Public IsFromTable As Boolean

Private sht As Object
Private ACAD As Object
Private mo As Object
Private math As New clsMath

Private Sub Class_Initialize()


Set sht = Sheets("橫斷面")
Set ACAD = MyACAD.acadDoc 'acadapp.ActiveDocument
Set mo = ACAD.ModelSpace

Xscale = 1
Yscale = 1

With sht

    For r = 2 To .Cells(.Rows.Count, "N").End(xlUp).Row
        collLineCD.Add .Cells(r, "N").Value
    Next

End With

'------------------------------

End Sub


Sub DrawYUNuseLine(ByVal sr As Integer, ByVal lr As Integer) 'plot the width of construction

'地界線

Dim Ldpt(2) As Double, lmpt(2) As Double, lupt(2) As Double
Dim rdpt(2) As Double, rmpt(2) As Double, rupt(2) As Double
Dim txtpt(2) As Double
Dim lineObj As Object
Dim collWidth As New Collection
Dim collMyCD As New Collection
Dim collYUNuse As New Collection
Dim collYUNuseWidth As New Collection

'On Error Resume Next

With sht

For r = sr To lr

    d = .Cells(r, 1)
    CD = .Cells(r, 3)
    
    For i = 1 To collLineCD.Count
    
        If CD = collLineCD.Item(i) Then
            
            collWidth.Add d * 1000 / Xscale
            collMyCD.Add CD
        
        End If
    
    Next
    
Next

For i = 1 To collMyCD.Count

    If collMyCD.Item(i) = "地界線" Then
    
        collYUNuse.Add collMyCD.Item(i)
        collYUNuseWidth.Add collWidth.Item(i)
    
    End If

Next

For i = 1 To collYUNuse.Count - 1
    
    lmpt(0) = BaseHeightPoint(0) + collYUNuseWidth.Item(1)
    rmpt(0) = BaseHeightPoint(0) + collYUNuseWidth.Item(2)
    lmpt(1) = BaseHeightPoint(1) + (step / 2 + 0.5) * 1000 / Yscale
    rmpt(1) = lmpt(1)

    Set lineObj = mo.AddLine(lmpt, rmpt)

Next

'Exit Sub

For i = 1 To collWidth.Count

    Ldpt(0) = BaseHeightPoint(0) + collWidth.Item(i)
    lupt(0) = BaseHeightPoint(0) + collWidth.Item(i)
    Ldpt(1) = BaseHeightPoint(1) + (step / 2 + 0.1) * 1000 / Yscale
    lupt(1) = BaseHeightPoint(1) + (step / 2 + 0.9) * 1000 / Yscale
    
    Set lineObj = mo.AddLine(Ldpt, lupt)
    'Update
    Set txtobj = PlotTextCenter(collMyCD.Item(i), lupt, 300 / Yscale, 1)

    If collMyCD.Item(i) = "中心樁" Then
    
        lineObj.Layer = "中心樁輔助線"
        txtobj.Layer = "中心樁輔助線"
        
    End If

Next
    
End With

End Sub

Sub DrawEGLine(ByVal sr As Integer, ByVal lr As Integer) 'ElevationGroundLine

Dim spt(2) As Double
Dim ept(2) As Double
Dim txtpt(2) As Double
Dim vertices() As Double
Dim retStart(2) As Double
Dim retEnd(2) As Double
Dim t As Integer

spt(0) = BaseHeightPoint(0)
spt(1) = BaseHeightPoint(1)

With sht
    
    For r = sr To lr
    
        CD = .Cells(r, 3)
        
        For i = 1 To collLineCD.Count
        
            If CD = collLineCD.Item(i) Then minus = minus + 1 '碰到地界線的時候EL=0
        
        Next
        
    Next
    
    t = lr - sr + 1 - minus

    ReDim vertices((t) * 3 - 1)
    
    myLocation = .Cells(sr, 4)
    
    For r = sr To lr
    
        d = .Cells(r, 1)
        CD = .Cells(r, 3)
        Y = .Cells(r, 2)
        
        For i = 1 To collLineCD.Count
        
            If CD = collLineCD.Item(i) Then GoTo jump
        
        Next
      
        ept(0) = spt(0) + d * 1000 / Xscale
        ept(1) = spt(1)
        
        txtpt(0) = ept(0)
        txtpt(1) = ept(1) + (Y - MidHeight) * 1000 / Yscale
        
        'Set txtobj = PlotTextCenter(cd, txtpt, 200 / Xscale, 1)
        
        vertices(0 + 3 * rr) = BaseHeightPoint(0) + d * 1000 / Xscale
        vertices(1 + 3 * rr) = BaseHeightPoint(1) + (Y - MidHeight) * 1000 / Yscale
        
        pt = InStr(1, .Cells(r, 4), ",")
        
        If pt <> 0 Then 'plot the open channel
        
            Dim obj As New clsUChannel
            Dim cpt(2) As Double
            
            PlanDiff = mid(.Cells(r, 4), 1, pt - 1) '計畫高落差
            ChannelType = mid(.Cells(r, 4), pt + 1) '渠道樣式
            
            cpt(0) = vertices(0 + 3 * rr)
            'cpt(1) = vertices(1 + 3 * rr) + Val(PlanDiff) * 1000
            cpt(1) = BaseHeightPoint(1) + (PlanDiff - MidHeight) * 1000 / Yscale
            
            obj.CenterPoint = cpt
            
            'Call MyACAD.InsertBlock(cpt, "test")
            obj.ReadDataFromXLS (ChannelType)
            obj.DrawMainStructure
            'obj.HsectionTitle (PlanDiff)
            
        End If
        
        Select Case r 'plot ground line
        
        Case sr
        
            retStart(0) = vertices(0 + 3 * rr)
            retStart(1) = vertices(1 + 3 * rr)
            
            Set o = MyACAD.InsertBlock(retStart, "LeftGL", PaperScale)
            
'            retEnd(0) = retStart(0) - 1500 / Xscale
'            retEnd(1) = retStart(1)
'
'            Call GroundLevel(retStart, retEnd)
            
        Case lr
        
            retStart(0) = vertices(0 + 3 * rr)
            retStart(1) = vertices(1 + 3 * rr)
            
            Set o = MyACAD.InsertBlock(retStart, "RightGL", PaperScale)
'
'            retEnd(0) = retStart(0) + 1500 / Xscale
'            retEnd(1) = retStart(1)
'
'            Call GroundLevel(retStart, retEnd)
            
        End Select
        
        rr = rr + 1
        
jump:
        
    Next
    
    'ACAD.ActiveLayer = ACAD.layers("橫斷面-原地面線")

    Set plineobj = MyACAD.AddPolyLine(vertices)
    
    'plineobj.Elevation = BaseHeightPoint(2) + TranLoc(Mylocation)
    
    'ACAD.ActiveLayer = ACAD.layers("0")

End With

End Sub

Sub GetCLHeight() '中心線高程採用前後高程的平均值

Dim collBlank As New Collection

With sht

    collBlank.Add 1
    
    For r = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
    
        If .Cells(r, 1) = "" Then collBlank.Add r
    
    Next

    collBlank.Add .Cells(Rows.Count, 1).End(xlUp).Row + 1

    For i = 1 To collBlank.Count - 1
    
        sr = collBlank.Item(i) + 1
        er = collBlank.Item(i + 1) - 1
        myloc = .Cells(sr, 4)
        
        For r = sr To er
        
            X = .Cells(r, 1)
            H = .Cells(r, 2)
            CD = .Cells(r, 3)
        
            If CD = "小給中心樁" Or CD = "小排中心樁" Or CD = "中心樁" Or CD = "CL" Then
            
                r1 = GetNonZeroRow(sr, er, r, 1)
                r2 = GetNonZeroRow(sr, er, r, 2)
                
                If r1 <> 0 And r2 <> 0 Then
                
                X1 = Abs(.Cells(r1, 1) - X)
                X2 = Abs(.Cells(r2, 1) - X)
            
                H1 = .Cells(r1, 2)
                H2 = .Cells(r2, 2)
                
                If H = 0 Then .Cells(r, 2) = (X1 * H2 + X2 * H1) / (X1 + X2)
                
                Else
                
                    MsgBox "樁號:" & myloc & "出現異常，等等需手動處理"
            
                End If
            
            End If
        
        Next
    
    Next
    
End With

End Sub

Function GetNonZeroRow(ByVal sr As Integer, ByVal er As Integer, ByVal r As Integer, ByVal mode As Byte)

'取該特定CD碼的前後高程不為0的平均值

With sht

    Select Case mode
    
    Case 1
    
        For tr = r - 1 To sr Step -1
            
            If .Cells(tr, 2) <> 0 Then
                GetNonZeroRow = tr
                Exit Function
            End If
        Next
    
    Case 2
    
        For tr = r + 1 To er
            
            If .Cells(tr, 2) <> 0 Then
                GetNonZeroRow = tr
                Exit Function
            End If
        Next
    
    End Select

End With

End Function

