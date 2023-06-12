Attribute VB_Name = "繪圖模組"

'Dim sset As AcadSelectionSet

Sub BasicSetting()

Dim txtStyle As AcadTextStyle
Dim txtStyles As AcadTextStyles
Dim lay As AcadLayer

Set txtStyles = ThisDrawing.textstyles

For Each txtStyle In txtStyles

    If txtStyle.Name = "工程用仿宋體" Then IsAdded = True

Next

If IsAdded = False Then

    Set txtStyle = txtStyles.Add("工程用仿宋體")
    
    On Error GoTo ERRORHANDLE
    
    txtStyle.fontFile = "C:\windows\fonts\SimSun.ttf"

End If

Set txtStyle = txtStyles("工程用仿宋體")

If ThisDrawing.activetextstyle.Name <> "工程用仿宋體" Then ThisDrawing.activetextstyle = txtStyle

If ThisDrawing.ActiveLayer.Name <> "0" Then ThisDrawing.ActiveLayer = ThisDrawing.Layers("0")

Set lay = ThisDrawing.Layers.Add("鋼筋層")
    lay.Color = acRed
    
Set lay = ThisDrawing.Layers.Add("標註層")
    lay.Color = acGreen
    
Set lay = ThisDrawing.Layers.Add("結構層")
    lay.Color = acWhite
    lay.LineWeight = acLnWt030

Set lay = ThisDrawing.Layers.Add("原地面線")
    lay.Color = acCyan
    
Set lay = ThisDrawing.Layers.Add("說明")
    lay.Color = acCyan

Set lay = ThisDrawing.Layers.Add("剖面圖說明")
    lay.LineWeight = acLnWt060
    lay.Color = acGreen
    
Set lay = ThisDrawing.Layers.Add("中心層")
    lay.Color = acRed
    
Set lay = ThisDrawing.Layers.Add("出圖圖框")
    lay.Color = acRed
    
Set lay = ThisDrawing.Layers.Add("出圖內框")
    lay.Color = acWhite
    
Set lay = ThisDrawing.Layers.Add("地盤高")
    lay.Color = acMagenta
        
Set lay = ThisDrawing.Layers.Add("計畫高")
    lay.Color = acRed
    lay.LineWeight = acLnWt035

Set lay = ThisDrawing.Layers.Add("左田高")
    lay.Color = acYellow
        
Set lay = ThisDrawing.Layers.Add("右田高")
    lay.Color = acCyan
    
Set lay = ThisDrawing.Layers.Add("鋼筋標註層")
    lay.Color = 140

Exit Sub

ERRORHANDLE:

txtStyle.fontFile = "C:\Windows\fonts\arial.ttf"

End Sub

Sub Hatch(ByVal obj As Object, ByVal Ratio As Double, Optional ByVal mode As Byte = 1)

Dim hatchobj As AcadHatch
Dim outerloop(0 To 0) As AcadEntity

Select Case mode

Case 1
    PatternName = "SOLID"
Case 2
    PatternName = "ANSI32"
    
End Select

Set hatchobj = ThisDrawing.ModelSpace.AddHatch(0, PatternName, True)

hatchobj.PatternScale = 1 / Ratio * 4

Set outerloop(0) = obj

hatchobj.AppendOuterLoop (outerloop)

End Sub

Function PlotRec(ByRef LeftLowerPoint() As Double, ByRef RightUpperPoint() As Double)

Dim vertices(5 * 3 - 1) As Double
Dim Rec 'As AcadPolyline

X1 = LeftLowerPoint(0): Y1 = LeftLowerPoint(1)
X2 = RightUpperPoint(0): Y2 = RightUpperPoint(1)

vertices(0) = X1: vertices(1) = Y1
vertices(3) = X1: vertices(4) = Y2
vertices(6) = X2: vertices(7) = Y2
vertices(9) = X2: vertices(10) = Y1
vertices(12) = X1: vertices(13) = Y1

Set PlotRec = AutoCAD.ActiveDocument.ModelSpace.AddPolyLine(vertices)

End Function

Function PlotText(ByVal s As String, ByRef txtpt() As Double, ByVal mode As Byte, ByVal ImportantMode As Byte, ByVal UserForm As UserForm)

Dim txtobj As AcadText

Ratio = UserForm.txtScale

Select Case ImportantMode

Case 1
    H = 150 / Ratio
Case 2
    H = 100 / Ratio
Case 3
    H = 80 / Ratio
    
End Select

Set txtobj = ThisDrawing.ModelSpace.AddText(s, txtpt, H)

Select Case mode

Case 1
    txtobj.Alignment = acAlignmentMiddleLeft
Case 2
    txtobj.Alignment = acAlignmentMiddleCenter
Case 3
    txtobj.Alignment = acAlignmentMiddleRight
    
End Select

txtobj.TextAlignmentPoint = txtpt

Set PlotText = txtobj

End Function

Function PlotTextCenter(ByVal s As String, ByVal txtpt, ByVal H As Double, ByVal mode As Byte, Optional circlemode As Boolean = False)

Dim txtobj As Object 'AcadText
Dim circleobj As Object 'AcadCircle
Dim MyACAD As New clsACAD


With MyACAD

    Set txtobj = MyACAD.AddText(s, txtpt, H, mode)
    
    Set PlotTextCenter = txtobj
    
    If circlemode = True Then Set circleobj = MyACAD.AddCircle(txtpt, H)

End With

End Function

Function PlotDim(ByRef pt1() As Double, ByRef PT2() As Double, ByVal UserForm As UserForm, ByVal mode As Byte, ByVal s As String, Optional specific As Byte = 1)

If s = 0 Then Exit Function

Dim lineObj, txtobj
Dim spt(2) As Double, ept(2) As Double
Dim dspt(2) As Double, dept(2) As Double
Dim txtpt(2) As Double
Dim ExtExtend As Double

With UserForm

    Unit = .cboUnit
    Ratio = .txtScale
    Multiple = GetMultiplemm(Unit)
    
End With

ExtOffset = 60 / Ratio
ExtExtend = 40 / Ratio
DimLine = 200 / Ratio

With ThisDrawing

    Select Case mode
        
    Case 1 '左
        
        spt(0) = pt1(0) - ExtOffset: spt(1) = pt1(1)
        ept(0) = spt(0) - DimLine: ept(1) = pt1(1)
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        spt(0) = PT2(0) - ExtOffset: spt(1) = PT2(1)
        ept(0) = spt(0) - DimLine: ept(1) = PT2(1)
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        spt(0) = pt1(0) - DimLine - ExtOffset + ExtExtend: spt(1) = pt1(1)
        ept(0) = PT2(0) - DimLine - ExtOffset + ExtExtend: ept(1) = PT2(1)
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        Call PlotDimOther(spt, ept, ExtExtend)
        
        txtpt(0) = (spt(0) + ept(0)) / 2 - ExtExtend * 4: txtpt(1) = (spt(1) + ept(1)) / 2
        Set txtobj = PlotText(s * Ratio / Multiple, txtpt, 2, 3, UserForm)
        
    Case 2 '上
    
        spt(0) = pt1(0): spt(1) = pt1(1) + ExtOffset
        ept(0) = spt(0): ept(1) = spt(1) + DimLine
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        spt(0) = PT2(0): spt(1) = PT2(1) + ExtOffset
        ept(0) = spt(0): ept(1) = spt(1) + DimLine
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        spt(0) = pt1(0): spt(1) = pt1(1) + DimLine + ExtOffset - ExtExtend
        ept(0) = PT2(0): ept(1) = PT2(1) + DimLine + ExtOffset - ExtExtend
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        Call PlotDimOther(spt, ept, ExtExtend)
             
        txtpt(0) = (spt(0) + ept(0)) / 2: txtpt(1) = (spt(1) + ept(1)) / 2 + ExtExtend * 3
        
        If specific = 2 Then txtpt(0) = txtpt(0) - ExtExtend * 2
        
        Set txtobj = PlotText(s * Ratio / Multiple, txtpt, 2, 3, UserForm)
        
    Case 3 '右
    
        spt(0) = pt1(0) + ExtOffset: spt(1) = pt1(1)
        ept(0) = spt(0) + DimLine: ept(1) = pt1(1)
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        spt(0) = PT2(0) + ExtOffset: spt(1) = PT2(1)
        ept(0) = spt(0) + DimLine: ept(1) = PT2(1)
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        spt(0) = pt1(0) + DimLine + ExtOffset - ExtExtend: spt(1) = pt1(1)
        ept(0) = PT2(0) + DimLine + ExtOffset - ExtExtend: ept(1) = PT2(1)
        
        Set lineObj = .ModelSpace.AddLine(spt, ept)
        
        Call PlotDimOther(spt, ept, ExtExtend)
        
        txtpt(0) = (spt(0) + ept(0)) / 2 + ExtExtend * 2: txtpt(1) = (spt(1) + ept(1)) / 2
        Set txtobj = PlotText(s * Ratio / Multiple, txtpt, 2, 3, UserForm)
        
    Case 4 '下
        
    End Select

End With

End Function

Sub PlotDimOther(ByRef spt() As Double, ByRef ept() As Double, ByVal ExtExtend As Double)

Dim dspt(2) As Double, dept(2) As Double

With ThisDrawing

        dspt(0) = spt(0) - ExtExtend: dspt(1) = spt(1) + ExtExtend
        dept(0) = spt(0) + ExtExtend: dept(1) = spt(1) - ExtExtend
        
        Set lineObj = .ModelSpace.AddLine(dspt, dept)
        
        dspt(0) = ept(0) - ExtExtend: dspt(1) = ept(1) + ExtExtend
        dept(0) = ept(0) + ExtExtend: dept(1) = ept(1) - ExtExtend
        
        Set lineObj = .ModelSpace.AddLine(dspt, dept)
        
End With
        
End Sub

Function ChooseDirection(ByVal spt As Variant, ByVal ept As Variant, ByVal dirpt As Variant)

X1 = spt(0)
Y1 = spt(1)

X2 = ept(0)
Y2 = ept(1)

Xd = dirpt(0)
Yd = dirpt(1)

If X1 = X2 Then
    mode = 1
ElseIf Y1 = Y2 Then
    mode = 2
Else
    ChooseDirection = 0
    Exit Function
End If

Select Case mode

Case 1

    If Y2 > Y1 Then
    
        If Xd > X1 Then
            ChooseDirection = 3 '右
        Else
            ChooseDirection = 1 '左
        End If
        
    Else 'Y1>Y2
    
        If Xd > X1 Then
            ChooseDirection = 7 '右
        Else
            ChooseDirection = 5 '左
        End If
        
    End If
    
Case 2

    If X2 > X1 Then
    
        If Yd < Y1 Then
            ChooseDirection = 4 '下
        Else
            ChooseDirection = 2 '上
        End If
        
    Else 'X1>X2
    
        If Yd < Y1 Then
            ChooseDirection = 8 '下
        Else
            ChooseDirection = 6 '上
        End If
        
    End If

End Select

End Function

Sub PlotArrow(ByVal X1 As Double, ByVal Y1 As Double, ByVal m As Double, ByVal arrowlength As Double, ByVal dire)

Dim vertices(4 * 3 - 1) As Double
Dim ept(2) As Double
Dim sret(2) As Double

With ThisDrawing

    sret(0) = X1: sret(1) = Y1

    vertices(0) = sret(0): vertices(1) = sret(1)
    ept(0) = X1 + arrowlength * 1 / (Sqr(m * m + 1)) * dire: ept(1) = Y1 + arrowlength * m / (Sqr(m * m + 1)) * dire

    Set lineObj = .ModelSpace.AddLine(sret, ept)

    lineObj.rotate sret, 1.9 * 3.14
    vertices(3) = lineObj.endpoint(0): vertices(4) = lineObj.endpoint(1)
    lineObj.Delete
    
    Set lineObj = .ModelSpace.AddLine(sret, ept)
    
    lineObj.rotate sret, 0.1 * 3.14
    vertices(6) = lineObj.endpoint(0): vertices(7) = lineObj.endpoint(1)
    lineObj.Delete
    
    vertices(9) = sret(0): vertices(10) = sret(1)
    
    Set plineobj = .ModelSpace.AddPolyLine(vertices)

    Call Hatch(plineobj, 1, 1)

End With

End Sub

Function CreateSSET(Optional mode As Byte = 0, Optional setname As String = "SS1")

Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim MyACAD As New clsACAD

Set ACAD = MyACAD.acaddoc '.ActiveDocument

With ACAD

    On Error Resume Next
        .SelectionSets(setname).Delete
    On Error GoTo 0
    
    Set sset = .SelectionSets.Add(setname)
    
    Select Case mode '變換方式
    
    Case 1: sset.SelectOnScreen
    
    Case 2:
    
    FilterType(0) = 0
    FilterData(0) = "Hatch"
    
    sset.SelectOnScreen FilterType, FilterData
    
    Case 3:
    
    FilterType(0) = 8
    FilterData(0) = "橫斷面樁"
    
    sset.SelectOnScreen FilterType, FilterData
    
    Case 4:
    
    FilterType(0) = 0
    FilterData(0) = "polyline"

    sset.SelectOnScreen FilterType, FilterData
    
    Case 5:
    
    FilterType(0) = 8
    FilterData(0) = "VPORT"

    sset.SelectOnScreen FilterType, FilterData
    
    
    End Select
    
    Set CreateSSET = sset
    'sset.Item(0).Delete

End With

End Function

Function NoteSeperate(ByVal s As String) As Variant

Dim collNote As New Collection
Dim NoteArr As Variant
IsCollected = False

For i = 1 To Len(s)
    
    ch = mid(s, i, 1)
    
    NoteString = NoteString & ch
    
    If ch = "、" Then
        NoteString = mid(NoteString, 1, Len(NoteString) - 1)
        collNote.Add NoteString
        NoteString = ""
        IsCollected = True
    End If
    
Next

collNote.Add NoteString

ReDim NoteArr(1 To collNote.Count)

For j = 1 To collNote.Count
    NoteArr(j) = collNote.Item(j)
Next

NoteSeperate = NoteArr

End Function


