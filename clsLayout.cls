VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private r As Integer
Private sr As Integer
Private lr As Integer
Private ACAD As Object
Private sht As Object
Private ps As Object
Private Count As Integer

Private collHead As New Collection
Private collContent As New Collection
Private collfinalWidth As Variant

Public X As Double
Public Y As Double
Public dx As Double
Public dy As Double
Public MyConfigName As String '繪圖機規劃
Public MyCanonicalMediaName As String '印表機圖紙
Public MyStyleSheet As String '出圖型式表

Private LayoutArr As Variant
Private MSLayoutArr As Variant
Private MSNumCol As Integer
Private MyACAD As New clsACAD

Private Sub Class_Initialize()

Dim layout As Object

Set ACAD = MyACAD.acaddoc

Set ps = ACAD.PaperSpace
Set mo = ACAD.ModelSpace
Set sht = Sheets("圖說")

ACAD.ActiveLayer = ACAD.Layers("FRAME")

sr = 3

MSNumCol = getMSNumCol '取得"圖號名稱"欄位

Call getSepLength '分配欄位寬度

'Call MyACAD.Createfontstyle

End Sub

Sub setLayoutDetail() '針對一次性的更新，也不再動圖框了

For Each layout In ACAD.Layouts

    layout.StyleSheet = MyStyleSheet

Next

End Sub

Sub getLayoutXLS()

With sht

    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    For c = 1 To 14
        collHead.Add .Cells(sr - 1, c).Value
    Next
    
    For r = sr To lr
    
        tmp = ""
    
        For c = 1 To 14
        
            If collHead(c) = "頁次" Then
                j = j + 1
                txt = "第" & j & "頁共" & lr - sr + 1 & "頁"
            Else
                txt = .Cells(r, c)
            End If
        
            tmp = tmp & "," & txt
            
        Next
        
        collContent.Add mid(tmp, 2)
        
    Next

End With

End Sub

Sub SortLayout(ByVal IsAutoCAD As Boolean)

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
        
        If IsAutoCAD Then
        
            If beforeID > afterID Then
                
                LayoutArr(j, 1) = afterID
                Set LayoutArr(j, 2) = afterObj
                LayoutArr(k, 1) = beforeID
                Set LayoutArr(k, 2) = beforeObj
                
            End If
        
        Else
        
            If beforeID < afterID Then
                
                LayoutArr(j, 1) = afterID
                Set LayoutArr(j, 2) = afterObj
                LayoutArr(k, 1) = beforeID
                Set LayoutArr(k, 2) = beforeObj
                
            End If

        End If
    
    Next

Next

End Sub

Sub FillInLayout()

With sht

i = 1

    For r = sr To lr

        psName = .Cells(r, MSNumCol) '.Cells(r, 3) & "," & i
        
        Set layout = LayoutArr(i, 2)
        ACAD.ActiveLayout = layout
        layout.ConfigName = MyConfigName
        layout.CanonicalMediaName = MyCanonicalMediaName
        layout.StyleSheet = MyStyleSheet
        layout.Name = psName
        
        Set ps = ACAD.PaperSpace
        Debug.Print ps.layout.Name
        Call CreateFrame_Main
        
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

Sub CreateFrame_Main()

'Call PlotFrameOuter
Call PlotFrame
'Call getSepLength
Call PlotFrameInner2

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

Sub ClearVport()

For Each layout In ACAD.Layouts

If Not layout.Name = "Model" Then layout.Delete

Next

End Sub

Private Sub PlotFrameOuter()

Dim txtpt(2) As Double
Dim vertices(5 * 2 - 1) As Double

vertices(0) = X - 20
vertices(1) = Y - 10

vertices(2) = X - 20
vertices(3) = Y + 277 + 10

vertices(4) = X + 390 + 10
vertices(5) = Y + 277 + 10

vertices(6) = X + 390 + 10
vertices(7) = Y - 10

vertices(8) = X - 20
vertices(9) = Y - 10

Set lwplineobj = ps.AddLightWeightPolyline(vertices)

lwplineobj.Layer = "出圖外框"

txtpt(0) = X + 300
txtpt(1) = Y - 20

Set txtobj = ps.AddText("出圖圈選範圍線(A3)    297*420", txtpt, 5)

txtobj.Color = 1 'acRed

End Sub

Private Sub PlotFrame()

Dim vertices(5 * 3 - 1) As Double
Dim plineobj As Object ' AcadPolyline

vertices(0) = X
vertices(1) = Y

vertices(3) = X
vertices(4) = Y + 277

vertices(6) = X + 390
vertices(7) = Y + 277

vertices(9) = X + 390
vertices(10) = Y

vertices(12) = X - 0.8
vertices(13) = Y

Set plineobj = ps.AddPolyLine(vertices)

plineobj.ConstantWidth = 1.6

End Sub

Sub getSepLength()

Dim collmyWidth As New Collection
Dim collfinalWidthtmp As New Collection

With sht

For c = 1 To MSNumCol + 1

    myWidth = .Columns(c).ColumnWidth

    collmyWidth.Add myWidth
    
    sumWidth = sumWidth + myWidth

Next

For i = 1 To collmyWidth.Count

    finalWidth = collmyWidth.Item(i) / sumWidth * 390

    collfinalWidthtmp.Add finalWidth

Next

Set collfinalWidth = collfinalWidthtmp

End With

End Sub

Private Sub PlotFrameInner2()

Dim vertices(3 * 2 - 1) As Double
Dim txtpt(2) As Double

Count = Count + 1

'arr = Array(60, 60, 50, 40, 22, 22, 22, 22, 20, 24, 20, 30) '分割長度(由左到右)

tmptxt = Split(collContent(Count), ",")

For i = 1 To 2

    vertices(0) = X
    vertices(1) = Y + 9.25 * i
    
    vertices(3) = X + 390
    vertices(4) = Y + 9.25 * i
    
    Set plineobj = ps.AddPolyLine(vertices)
    
    plineobj.ConstantWidth = 0.5
    
    Xt = 0
    
    For j = 0 To collfinalWidth.Count - 1

        Xt = Xt + collfinalWidth.Item(j + 1) 'arr(j)
    
        Xtt = Xt - 0.5 * collfinalWidth.Item(j + 1) 'arr(j)
        Yt = Y + 9.25 * i - 0.5 * 9.25
        
        txtpt(0) = Xtt
        txtpt(1) = Yt
        
        If i = 1 Then
            txtprompt = tmptxt(j)
        Else
            txtprompt = collHead(j + 1)
        End If
        
        Set txtobj = ps.AddText(txtprompt, txtpt, 3.5) '文字高度3.5
        txtobj.Alignment = 10 'acAlignmentMiddleCenter
        txtobj.TextAlignmentPoint = txtpt
        
        Call MyACAD.FitTextWidth(txtobj, collfinalWidth.Item(j + 1) * 0.8)
        
    Next
    
Next

Xorigin = X

For j = 0 To collfinalWidth.Count - 1 - 1 'UBound(arr) - 1

        Xorigin = Xorigin + collfinalWidth.Item(j + 1) ' arr(j)
        
        vertices(0) = Xorigin
        vertices(1) = Y
        vertices(3) = Xorigin
        vertices(4) = Y + 18.5
        Set plineobj = ps.AddPolyLine(vertices)
        plineobj.ConstantWidth = 0.5

Next

End Sub

Private Sub PlotFrameInner()

Dim vertices(3 * 2 - 1) As Double
Dim txtpt(2) As Double
Dim txtArr() As Variant

With sht

Lc = 11 '.Cells(1, 1).End(xlToRight).Column

ReDim txtArr(1 To Lc, 1 To 2)

    For c = 1 To Lc
    
        If .Cells(r, c) <> "" Then txtArr(c, 1) = .Cells(r, c)
        txtArr(c, 2) = .Cells(1, c)
        
    Next

End With

txtArr(Lc, 1) = "第" & r - 1 & "頁共" & lr - 1 & "頁"

arr = Array(60, 90, 70, 20, 20, 20, 20, 20, 24, 16, 30) '分割長度(由左到右)

For i = 1 To 2

    vertices(0) = X
    vertices(1) = Y + 9.25 * i
    
    vertices(3) = X + 390
    vertices(4) = Y + 9.25 * i
    
    Set plineobj = ps.AddPolyLine(vertices)
    
    plineobj.ConstantWidth = 0.5
    
    Xt = vertices(0)
    
    For j = 0 To UBound(arr)
    
        Xt = Xt + arr(j)
    
        Xtt = Xt - 0.5 * arr(j)
    
        Yt = vertices(1) - 0.5 * 9.25
        
        txtpt(0) = Xtt
        txtpt(1) = Yt
        
        Set txtobj = ps.AddText(txtArr(j + 1, i), txtpt, 3.5) '文字高度3.5
        
        If Len(txtobj.TextString) >= 14 Then txtobj.height = 3
        
        txtobj.Alignment = 10 'acAlignmentMiddleCenter
        txtobj.TextAlignmentPoint = txtpt
        
    Next
    
Next

Xorigin = X

For j = 0 To UBound(arr) - 1

        Xorigin = Xorigin + arr(j)
        
        vertices(0) = Xorigin
        vertices(1) = Y
        vertices(3) = Xorigin
        vertices(4) = Y + 18.5
        Set plineobj = ps.AddPolyLine(vertices)
        plineobj.ConstantWidth = 0.5

Next

End Sub

Sub CreateMSVport()

Dim Ldpt(2) As Double
Dim rupt(2) As Double
Dim txtpt(2) As Double

ret = ACAD.Utility.GetPoint(, "請選擇圖框產生的起始點")

X_MS = ret(0): Y_MS = ret(1)

ACAD.ActiveLayer = ACAD.Layers("0")

With sht

    lr = .Cells(Rows.Count, 1).End(xlUp).Row

    For r = sr To lr
        
        MSNum = .Cells(r, MSNumCol)
        MSVportName = MSNum & ":" & .Cells(r, 3)
        MSUnit = .Cells(r, 12)
        MSScale = .Cells(r, 13)
        MSnote = .Cells(r, 14)
        
        txtheight = 20
        If Len(MSVportName) > 13 Then txtheight = 15
        
        If r > 2 Then X_MS = X_MS + 500
        
        If ii Mod 10 = 0 And ii <> 0 Then
            
            'Y_MS = Y_MS - 500: X_MS = ret(0)
            
        End If
        
        Ldpt(0) = X_MS: Ldpt(1) = Y_MS
        rupt(0) = X_MS + dx: rupt(1) = Y_MS + dy
    
        Set plineobj = MyACAD.PlotRec(Ldpt, rupt)
        plineobj.Layer = "VPORT"
    
        txtpt(0) = (Ldpt(0) + rupt(0)) / 2
        txtpt(1) = (Ldpt(1) + rupt(1)) / 2
    
        Set txtobj = PlotTextCenter(MSVportName, txtpt, txtheight, 2)
        txtobj.Layer = "VPORT"
        
        txtpt(0) = Ldpt(0) + 5: txtpt(1) = Ldpt(1) + 10
        If MSUnit <> "" Then
            Set txtobj = PlotTextCenter("UNIT=" & MSUnit, txtpt, 5, 1) 'Unit
            Call MyACAD.AddTextBox(txtobj, 2)
        End If
        txtpt(0) = Ldpt(0) + 5: txtpt(1) = Ldpt(1) + 20
        If MSScale <> "" Then
            Set txtobj = PlotTextCenter("SCALE=" & MSScale, txtpt, 5, 1) 'Scale
            Call MyACAD.AddTextBox(txtobj, 2)
        End If
        
        txtpt(0) = Ldpt(0) + 20: txtpt(1) = Ldpt(1) + 240
        'Set txtobj = PlotTextCenter(MSNum, txtpt, 10, 2, True)
        
        txtpt(0) = Ldpt(0) + 20: txtpt(1) = Ldpt(1) + 220
        'Set txtobj = PlotTextCenter(MSNote, txtpt, 10, 1) '圖標
        'txtobj.StyleName = "工程用仿宋體(直)"
        'txtobj.Rotate txtpt, -3.14 / 2
        
        ii = ii + 1
    
    Next

End With

End Sub

Sub CollectMSLayout()

Dim sset As Object ' AcadSelectionSet

ReDim MSLayoutArr(1 To lr - sr + 1, 1 To 4)

Dim collCenter As New Collection
Dim collScale As New Collection
Dim collLayoutName As New Collection
Dim collChName As New Collection

Set sset = CreateSSET(5, "SS1")

For Each Item In sset
    
    If Item.objectname = "AcDbText" Then
    
        s = Item.TextString
    
        Debug.Print Item.TextString
        pt = Item.TextAlignmentPoint
        
        collCenter.Add pt
        
        dot = InStr(1, s, ":")
        
        layoutname = mid(s, 1, dot - 1)
        chname = mid(s, dot + 1)
        
        collLayoutName.Add layoutname
        collChName.Add chname

    Else
    
        RecArr = Item.coordinates
    
        l = (RecArr(6) - RecArr(0))
        collScale.Add dx / l
        
    End If

Next

For i = 1 To collScale.Count

    MSLayoutArr(i, 1) = collLayoutName.Item(i)
    MSLayoutArr(i, 2) = collCenter.Item(i)
    MSLayoutArr(i, 3) = collScale.Item(i)
    MSLayoutArr(i, 4) = collChName.Item(i)

Next

End Sub

Function getMSNumCol()

With sht

    For c = 1 To 100
    
        If .Cells(2, c) = "圖號" Then
            getMSNumCol = c
            Exit Function
        End If
    
    Next

End With

End Function


Private Sub Class_Terminate()

'Update

End Sub
