VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLayout2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TODO:
'1.collectVport
'2.createLayout
'3.setLayout

Public MyConfigName As String '繪圖機規劃
Public MyCanonicalMediaName As String '印表機圖紙
Public MyStyleSheet As String '出圖型式表

Private shtLayout As Object
Private CAD As New clsACAD
Private sr As Integer
Private lr As Integer

Private LayoutArr As Variant
Private MSLayoutArr As Variant

Private Sub Class_Initialize()

Set shtLayout = Sheets("圖說")

sr = 3

With shtLayout

    lr = .Cells(2, 10).End(xlDown).Row

End With

End Sub

Sub CreateMSVport()

Dim pt_insert(2) As Double

c = CInt(InputBox("請輸入[圖號]欄位:", , 10))

ret = CAD.GetPoint("請選擇圖框產生的起始點")
pt_insert(0) = ret(0)
pt_insert(1) = ret(1)

With shtLayout
    
    'lr = .Cells(2, c).End(xlDown).Row
    
    For r = sr To lr
        
        If r <> sr Then pt_insert(0) = pt_insert(0) + 500
        pt_insert(1) = pt_insert(1)
        
        MSnote = .Cells(r, c) & ":" & .Cells(r, 3)
        
        Set Vport = CAD.InsertBlock(pt_insert, "繪圖區範圍(YLIA)", 1, Array(MSnote))
        
        Vport.Layer = "VPORT"
        
    Next

End With

End Sub

Sub getMSVport()

Set MSVports = CAD.CreateSSET("SS1", "8", "VPORT")

For Each MSVport In MSVports

    Set o = CAD.getBlockAttr(MSVport)
    
    Call CAD.GetBoundingBox(o, MinX, MinY, MaxX, MaxY)
    
    Call CAD.AddPointCO((MinX + MaxX) / 2, (MinY + MaxY) / 2)
    
Next

End Sub

Sub SortLayout(ByVal IsAutoCAD As Boolean)

Set ACAD = CAD.acadDoc

Dim layout As Object 'AcadLayout
ReDim LayoutArr(1 To lr - sr + 1, 1 To 2)

For Each layout In ACAD.Layouts

    If layout.Name <> "Model" Then c = c + 1

Next

For i = 2 To lr - sr + 1

    Set layout = ACAD.Layouts.Add(i & "-" & Round(Rnd(), 3))
    
Next

i = 1

For Each layout In ACAD.Layouts

    If Not layout.Name = "Model" Then
    
        LayoutArr(i, 1) = CStr(layout.objectid)
        Set LayoutArr(i, 2) = layout
    
        i = i + 1
    
    End If
    
Next

For j = 1 To UBound(LayoutArr, 1)

    For k = j To UBound(LayoutArr, 1)
    
        beforeID = LayoutArr(j, 1)
        Set beforeObj = LayoutArr(j, 2)
        afterID = LayoutArr(k, 1)
        Set afterObj = LayoutArr(k, 2)

        If beforeID > afterID Then

            LayoutArr(j, 1) = afterID
            Set LayoutArr(j, 2) = afterObj
            LayoutArr(k, 1) = beforeID
            Set LayoutArr(k, 2) = beforeObj

        End If

    Next

Next

End Sub

Sub ClearVport()

Set ACAD = CAD.acadDoc

For Each layout In ACAD.Layouts

If Not layout.Name = "Model" Then layout.Delete

Next

End Sub

Sub AddVport()

Dim sset As Object 'AcadSelectionSet
Dim vpt As Object 'AcadPViewport
Dim Center(2) As Double

Center(0) = 195.8
Center(1) = 152.75

Set vpt = ACAD.PaperSpace.AddPViewport(Center, dx, dy)
vpt.Layer = "VPORT"

vpt.ViewportOn = True
vpt.GridOn = True

End Sub

Sub FillInLayout()

Set ACAD = CAD.acadDoc

With shtLayout

i = 1

    For r = sr To lr

        psName = .Cells(r, 10) '.Cells(r, 3) & "," & i
        
        Set layout = LayoutArr(i, 2)
        ACAD.ActiveLayout = layout
        layout.ConfigName = MyConfigName
        layout.CanonicalMediaName = MyCanonicalMediaName
        layout.StyleSheet = MyStyleSheet
        layout.Name = psName
        
        Set ps = ACAD.PaperSpace
        Debug.Print ps.layout.Name
        'Call CreateFrame_Main
        
        For j = 1 To UBound(MSLayoutArr)
        
            If MSLayoutArr(j, 1) = psName Then
            
                Call AddVport
                
                Dim s As String
                
                pt = MSLayoutArr(j, 2)
                
                s = pt(0) & "," & pt(1)
                
                ACAD.SendCommand "Mspace" & vbCr & "ZOOM" & vbCr & "C" & vbCr & s & vbCr & vbCr
                
                ACAD.ActivePViewport.CustomScale = MSLayoutArr(j, 3)
                
                ACAD.SendCommand "-VPORTS" & vbCr & "L" & vbCr & "ON" & vbCr & "L" & vbCr & vbCr
                            
            
            End If
        
        Next
        
        i = i + 1
        
    Next

End With

End Sub
