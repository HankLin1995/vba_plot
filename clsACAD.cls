VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsACAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mo As Object
Private pa As Object
Private math As New clsMath
Private myFunc As New clsMyfunction
Public acadDoc As Object
Public CADVer As String

Private Sub Class_Initialize()

If Sheets("總表").optAutoCAD = True Then

    strCAD = "AutoCAD.application"
    CADVer = "AUTOCAD"
    
ElseIf Sheets("總表").optZWCAD = True Then
    
    strCAD = "ZWCAD.Application"
    CADVer = "ZWCAD"
    
ElseIf Sheets("總表").optICAD = True Then

    strCAD = "ICAD.application"
    CADVer = "ICAD"

End If

Call CADInit(strCAD)

End Sub

Private Sub CADInit(ByVal strCAD As String)

On Error Resume Next

Set acadApp = GetObject(, strCAD) '查看安裝
If Err <> 0 Then Set acadApp = CreateObject(strCAD)
acadApp.Visible = True

On Error GoTo 0

Set mo = acadApp.ActiveDocument.ModelSpace
Set pa = acadApp.ActiveDocument.PaperSpace
Set acadDoc = acadApp.ActiveDocument

End Sub

'=============transform ICAD about point and points======================

Function tranPoint(ByVal CADpt)

If CADVer <> "ICAD" Then tranPoint = CADpt: Exit Function

Set tranPoint = Library.CreatePoint(CADpt(0), CADpt(1), CADpt(2))

End Function

Function tranIPoint(ByVal ICADpt)

If CADVer <> "ICAD" Then tranIPoint = ICADpt: Exit Function

Dim tmp(2) As Double
tmp(0) = ICADpt.X
tmp(1) = ICADpt.Y
tmp(2) = ICADpt.Z

tranIPoint = tmp

End Function

Function tranPoints(ByVal vertices, Optional cnt As Byte = 3)

If CADVer <> "ICAD" Then tranPoints = vertices: Exit Function

If vertices(2) = 0 Then cnt = 3

Dim myPline, myPoints, pt
    
Set myPoints = Library.CreatePoints

If cnt = 2 Then

    For i = 0 To UBound(vertices) Step cnt
    
        Set pt = Library.CreatePoint(vertices(i), vertices(i + 1))
    
        myPoints.Add
        myPoints(myPoints.Count - 1).X = pt.X
        myPoints(myPoints.Count - 1).Y = pt.Y
    
    Next

Else

    For i = 0 To UBound(vertices) Step cnt
    
        Set pt = Library.CreatePoint(vertices(i), vertices(i + 1), vertices(i + 2))
    
        myPoints.Add
        myPoints(myPoints.Count - 1).X = pt.X
        myPoints(myPoints.Count - 1).Y = pt.Y
        myPoints(myPoints.Count - 1).Z = pt.Z
    
    Next

End If

Set tranPoints = myPoints

End Function

Function tranIPoints(ByVal myPoints)

If CADVer <> "ICAD" Then tranIPoints = myPoints: Exit Function

Dim vertices()
ReDim vertices(myPoints.Count * 3 - 1)

For Each it In myPoints

vertices(0 + j) = it.X
vertices(0 + j + 1) = it.Y
vertices(0 + j + 2) = it.Z

j = j + 3

Next

tranIPoints = vertices

End Function

'=============Basic function to CAD object======================

Function AddPoint(pt) As Object

If CADVer = "ICAD" Then
    Set AddPoint = mo.AddPointEntity(tranPoint(pt))
Else
    Set AddPoint = mo.AddPoint(tranPoint(pt))
End If

End Function

Function AddCircle(cpt, r) As Object

Set AddCircle = mo.AddCircle(tranPoint(cpt), r)

End Function

Function AddLine(spt, ept) As Object

Set AddLine = mo.AddLine(tranPoint(spt), tranPoint(ept))

End Function

Function AddXLine(spt, ept) As Object

Set AddXLine = mo.AddXLine(tranPoint(spt), tranPoint(ept))

End Function

Function AddLineCO(X1, Y1, X2, Y2) As Object

Dim spt(2) As Double
Dim ept(2) As Double
spt(0) = X1: spt(1) = Y1
ept(0) = X2: ept(1) = Y2

Set AddLineCO = AddLine(spt, ept)

End Function

Function AddPolyLine(vertices) As Object

Set AddPolyLine = mo.AddPolyLine(tranPoints(vertices))

End Function

Function Add3dPoly(vertices) '20210604 new

Set Add3dPoly = mo.Add3dPoly(tranPoints(vertices))

End Function

Function AddLWPolyLine(vertices) As Object

Set AddLWPolyLine = mo.AddLightWeightPolyline(tranPoints(vertices, 2))

End Function

Function AddArc(ByVal Center, Radius As Double, StartAngle As Double, EndAngle As Double)

Set AddArc = mo.AddArc(tranPoint(Center), Radius, StartAngle, EndAngle)

End Function

Function AddText(ByVal myText As String, ByVal txtpt, ByVal txtheight As Double, Optional alignmode As Byte = 1)

    Dim mtextObj As Object 'IntelliCAD.Text
    Dim insPt
    
    insPt = txtpt
    Set txtobj = mo.AddText(myText, tranPoint(insPt), txtheight)

    If CADVer = "ICAD" Then
    
        Select Case alignmode
        
        Case 1
            txtobj.HorizontalAlignment = 0  ' acAlignmentMiddleLeft
        Case 2
            txtobj.HorizontalAlignment = 4  ' acAlignmentMiddleCenter
        Case 3
            txtobj.HorizontalAlignment = 2 ' acAlignmentMiddleRight
        
        End Select
    
    Else
    
        Select Case alignmode
    
        Case 1
            txtobj.Alignment = 9 ' acAlignmentMiddleLeft
        Case 2
            txtobj.Alignment = 10 ' acAlignmentMiddleCenter
        Case 3
            txtobj.Alignment = 11 ' acAlignmentMiddleRight
            
        End Select
    
    End If
    
    txtobj.TextAlignmentPoint = tranPoint(insPt)

    Set AddText = txtobj

End Function

'=========rely on basic function===================

Function AddCircleCO(X, Y, r) As Object

Dim cpt(2) As Double: cpt(0) = X: cpt(1) = Y
Set AddCircleCO = AddCircle(cpt, r)

End Function

Function AddPointCO(X, Y) As Object

Dim pt(2) As Double: pt(0) = X: pt(1) = Y
Set AddPointCO = AddPoint(pt)

End Function

Function PlotRec(ByRef LeftLowerPoint() As Double, ByRef RightUpperPoint() As Double)

Dim vertices(5 * 3 - 1) As Double

X1 = LeftLowerPoint(0): Y1 = LeftLowerPoint(1)
X2 = RightUpperPoint(0): Y2 = RightUpperPoint(1)

vertices(0) = X1: vertices(1) = Y1
vertices(3) = X1: vertices(4) = Y2
vertices(6) = X2: vertices(7) = Y2
vertices(9) = X2: vertices(10) = Y1
vertices(12) = X1: vertices(13) = Y1

Set PlotRec = AddPolyLine(vertices)

End Function

Function PlotRecFillet(ByRef LeftLowerPoint() As Double, ByRef RightUpperPoint() As Double, ByVal r As Double)

Dim vertices(9 * 3 - 1) As Double

X1 = LeftLowerPoint(0): Y1 = LeftLowerPoint(1)
X2 = RightUpperPoint(0): Y2 = RightUpperPoint(1)

vertices(0) = X1: vertices(1) = Y1 + r
vertices(3) = X1 + r: vertices(4) = Y1
vertices(6) = X2 - r: vertices(7) = Y1
vertices(9) = X2: vertices(10) = Y1 + r
vertices(12) = X2: vertices(13) = Y2 - r
vertices(15) = X2 - r: vertices(16) = Y2
vertices(18) = X1 + r: vertices(19) = Y2
vertices(21) = X1: vertices(22) = Y2 - r
vertices(24) = X1: vertices(25) = Y1 + r

Set tmp = AddPolyLine(vertices)

coe = 0.35

tmp.SetBulge 0, coe
tmp.SetBulge 2, coe
tmp.SetBulge 4, coe
tmp.SetBulge 6, coe

Set PlotRecFillet = tmp

End Function

Function AddPolyLineCoor(ByVal CenterPoint, ByVal BasicParameterArr, ByVal ControlCalculation) As Object

Dim vertices() As Double
Dim spt(2) As Double
ReDim vertices(3 * (UBound(BasicParameterArr) + 2) - 1) '頭尾各加一

spt(0) = CenterPoint(0) 'tmp
spt(1) = CenterPoint(1) 'tmp

vertices(0) = spt(0)
vertices(1) = spt(1)

For i = 0 To UBound(BasicParameterArr)
    
    Calc = ControlCalculation(i)
    Shape = BasicParameterArr(i)
    
    Select Case Calc
    
    Case "X": spt(0) = spt(0) + Shape
    Case "Y": spt(1) = spt(1) + Shape
    Case Else:
        tmp = Split(Shape, ":")
        spt(0) = spt(0) + tmp(0)
        spt(1) = spt(1) + tmp(1)
    End Select

    vertices(3 + i * 3) = spt(0)
    vertices(4 + i * 3) = spt(1)
    
Next

Set AddPolyLineCoor = AddPolyLine(vertices)

End Function

Function FitTextWidth(ByVal txtobj As Object, ByVal FitWidth As Double)

Set entobj = txtobj
Call GetBoundingBox(entobj, X1, Y1, X2, Y2)

myWidth = X2 - X1

Do Until myWidth <= FitWidth

    txtobj.height = txtobj.height - 0.1
    Call GetBoundingBox(entobj, X1, Y1, X2, Y2)
    myWidth = X2 - X1

Loop

End Function

Function AddCText(ByVal s As String, ByRef txtpt() As Double, ByVal txtheight As Double)

Dim txtobj As Object 'AcadText
Dim circleobj As Object 'AcadCircle

With mo

    Set txtobj = AddText(s, txtpt, txtheight, 2)

    Set AddCText = txtobj
    
   Call GetBoundingBox(txtobj, MinX, MinY, MaxX, MaxY)
    
    r = Sqr((MaxX - MinX) ^ 2 + (MaxY - MinY) ^ 2) / 2
    
    Set circleobj = AddCircle(txtpt, r)
    
End With

End Function

Function AddTextBox(ByVal txtobj As Object, Optional mode As Byte = 1)

Dim Ldpt(2) As Double
Dim rupt(2) As Double

Call GetBoundingBox(txtobj, MinX, MinY, MaxX, MaxY)

r = Sqr((MaxX - MaxX) ^ 2 + (MaxY - MinY) ^ 2)

Ldpt(0) = MinX - 0.3 * r: Ldpt(1) = MinY - 0.3 * r
rupt(0) = MaxX + 0.3 * r: rupt(1) = MaxY + 0.3 * r

If mode = 1 Then
 
Set AddTextBox = PlotRecFillet(Ldpt, rupt, 0.4 * r)

Else

Set AddTextBox = PlotRec(Ldpt, rupt)

End If

End Function

'=========about utility================================

Function GetPoint(ByVal prompt As String)

If CADVer = "ICAD" Then

    Set o = IntelliCAD.ActiveDocument.Utility.GetPoint(, prompt)
    GetPoint = tranIPoint(o)

Else
    GetPoint = acadDoc.Utility.GetPoint(, prompt)
End If

End Function

Function GetString(ByVal prompt As String)

GetString = acadDoc.Utility.GetString(False, prompt & vbNewLine)

End Function

Function GetBoundingBox(ByVal entobj, MinX, MinY, MaxX, MaxY)

If CADVer = "ICAD" Then Call GetBoundingBox_ICAD(entobj, MinX, MinY, MaxX, MaxY): Exit Function

Call entobj.GetBoundingBox(Min, Max)

MinX = Min(0): MinY = Min(1)
MaxX = Max(0): MaxY = Max(1)

End Function

Function GetBoundingBox_ICAD(ByVal entobj, MinX, MinY, MaxX, MaxY)

Dim Min As Object
Dim Max As Object

Call entobj.GetBoundingBox(Min, Max)

MinX = Min.X: MinY = Min.Y
MaxX = Max.X: MaxY = Max.Y

End Function


'===============test not passed===============================

Function ReverseLine(ByVal tmpLine As Object) '轉換過程中雖然座標已經轉換但還是無法

Dim vertices() As Double
Dim collX As New Collection
Dim collY As New Collection
Dim collZ As New Collection

If CADVer = "ICAD" Then
    pts = tranIPoints(tmpLine.coordinates)
Else
    pts = tmpLine.coordinates
End If

co = 3

If TypeName(tmpLine) Like "*LWPolyline" Then co = 2

ReDim vertices(UBound(pts))

For i = 0 To UBound(pts) Step co

    collX.Add pts(i)
    collY.Add pts(i + 1)
    
    If co = 3 Then collZ.Add pts(i + 2)

Next

For j = collX.Count To 1 Step -1

    vertices(k) = collX(j)
    vertices(k + 1) = collY(j)
    
    If co = 3 Then vertices(k + 2) = collZ(j)
    
    k = k + co

Next

If co = 2 Then
    Set ReverseLine = AddLWPolyLine(vertices)
Else
    Set ReverseLine = AddPolyLine(vertices)
End If
tmpLine.Delete

End Function

Sub Createfontstyle() '這部分可以用DWT方法解決(先建立好水利會模板)

Dim txtStyle 'As AcadTextStyle
Dim txtStyles ' As AcadTextStyles

Set txtStyles = acadDoc.textstyles
'
'For Each txtStyle In txtStyles
'
'    If txtStyle.Name = "工程用仿宋體" Then IsAdded = True
'
'Next
'
'If IsAdded = False Then
'
'    Set txtStyle = txtStyles.Add("工程用仿宋體")
'
'    On Error GoTo ERRORHANDLE
'
'    txtStyle.fontFile = "C:\windows\fonts\SimSun.ttf"
'
'End If

acadDoc.activetextstyle = acadDoc.textstyles("YLIA")

Exit Sub

ERRORHANDLE:

txtStyle.fontFile = "C:\Windows\fonts\arial.ttf"

acadDoc.activetextstyle = acadDoc.textstyles("YLIA")

End Sub

Function AddMixText(ByVal s As String, ByVal txtpt, ByVal txtheight As Double, _
                    Optional alignmode As Byte = 1, Optional dashmode As Integer = 0)

Dim txtobj As Object 'AcadText
Dim circleobj As Object 'AcadCircle

Set AddMixText = AddText(s, txtpt, txtheight, alignmode)

With mo

    Dim vertices(5) As Double

    If dashmode = 0 Then Exit Function
    
    Call GetBoundingBox(AddMixText, MinX, MinY, MaxX, MaxY)
    
    For i = 1 To dashmode
    
        vertices(0) = MinX
        vertices(1) = MinY - txtheight / 2 - txtheight / 2 * 0.5 * i
        
        vertices(3) = MaxX
        vertices(4) = vertices(1)
        
        Set plineobj = AddPolyLine(vertices)
        
        '******整體寬度尚未解決******
        
        If CADVer <> "ICAD" Then plineobj.ConstantWidth = txtheight / 10
    
    Next

End With

End Function

Sub Hatch(ByVal obj As Object, ByVal Ratio As Double, ByVal ptrName As String)

'usually use "SOLID","ANSI32","AR-CONC"

If CADVer = "ICAD" Then Exit Sub

Dim hatchobj As Object 'AcadHatch
Dim outerloop(0 To 0) As Object 'AcadEntity

PatternName = ptrName

Set hatchobj = mo.AddHatch(0, PatternName, True)
'Set hatchobj = mo.AddHatch(vicHatchPatternTypePreDefined, PatternName, True)

hatchobj.PatternScale = 1 / Ratio * 4

Set outerloop(0) = obj

On Error Resume Next

hatchobj.AppendOuterLoop (outerloop)

End Sub

Function CreateSSET(Optional ByVal sname As String = "SS1", Optional ByVal ftypetmp As Variant = "", Optional ByVal fdatatmp As Variant = "")

'****PorgeCAD中似乎沒辦法做到過濾的條件?*****

'0:object type
'2:object name
'8:layer name
'62:color number(0 to 256)

Dim FilterType() As Integer
Dim FilterData() As Variant

On Error Resume Next: acadDoc.SelectionSets(sname).Delete: On Error GoTo 0

Set sset = acadDoc.SelectionSets.Add(sname)

If ftypetmp = "" Then

    sset.SelectOnScreen
    
Else

    ft = Split(ftypetmp, ",")
    fd = Split(fdatatmp, ",")
    
    ReDim FilterType(0 To UBound(ft))
    ReDim FilterData(0 To UBound(fd))
    
    For i = 0 To UBound(ft)
    
        FilterType(i) = ft(i)
        FilterData(i) = fd(i)
        
    Next
    
    sset.SelectOnScreen FilterType, FilterData
    
End If

Set CreateSSET = sset

End Function

Sub SetXdataToObj(ByVal obj As Object, ByVal xdata As String) '提供綁入字串使用

If CADVer = "ICAD" Then Exit Sub  '中心線樁無法提供
'*****用Handle處理?!***

Dim DataType(0 To 1) As Integer
Dim Data(0 To 1) As Variant

DataType(0) = 1001: Data(0) = xdata
DataType(1) = 1000: Data(1) = ""

obj.SetXData DataType, Data

End Sub



'****不實用****

Sub BeGroup(ByVal entArr As Variant)

Set GO = acadDoc.groups.Add(acadDoc.groups.Count + 1)

If TypeName(entArr) = "Collection" Then

    Dim tmpArr() As Object

    ReDim tmpArr(0 To entArr.Count - 1)

    For i = 1 To entArr.Count
        
        Set tmpArr(i - 1) = entArr.Item(i)
        
    Next

    GO.AppendItems tmpArr

    Exit Sub

End If

GO.AppendItems entArr

End Sub

'****直接用內建FUNCTION****

Function PlotArrow(ByVal spt, ByVal ept, Optional ByVal size As Double = 6, Optional ByVal rotAng As Double = 15)

Dim cpt(2) As Double
Dim vertices() As Double

fi = math.getAz(spt(0), spt(1), ept(0), ept(1))
thida = 90 - fi

dx = size
dy = size * Atn(math.deg2rad(rotAng))

cpt(0) = spt(0) + dx
cpt(1) = spt(1) + dy

shparr = Array(-dx & ":" & -dy, dx & ":" & -dy, -dx & ":" & dy, math.getLength(spt, ept))
calcarr = Array("", "", "", "X")

Set PlotArrow = AddPolyLineCoor(cpt, shparr, calcarr)
PlotArrow.rotate spt, math.deg2rad(thida)

'Exit Function

shparr = Array(-dx & ":" & -dy, dx & ":" & -dy, 2 * dy)
calcarr = Array("", "", "Y")

Set tmp = AddPolyLineCoor(cpt, shparr, calcarr)
tmp.rotate spt, math.deg2rad(thida)
Call Hatch(tmp, 1, "SOLID")

End Function

'****直接用內建FUNCTION****

Sub DimRotated(ByVal spt, ByVal ept, ByVal Dimtxtheight As Double, ByVal DimExtend As Double, _
               Optional ByVal DimArrowSize As Double = 6, Optional ByVal DimContent As String = "", _
               Optional DimScale As String = "m", Optional ByVal IsGroup As Boolean = False, Optional dx As Double = 0)

Dim mpt(2) As Double
Dim upt(2) As Double
Dim dpt(2) As Double

fi = math.getAz(spt(0), spt(1), ept(0), ept(1))

mpt(0) = (spt(0) + ept(0)) / 2
mpt(1) = (spt(1) + ept(1)) / 2

mylen = math.getLength(spt, ept)

If DimArrowSize > mylen / 2 * 0.8 Then DimArrowSize = mylen * 0.4

Set arrow1 = PlotArrow(spt, mpt, DimArrowSize) '左箭頭
Set arrow2 = PlotArrow(ept, mpt, DimArrowSize) '右箭頭

upt(0) = spt(0): upt(1) = spt(1) + DimExtend
dpt(0) = spt(0): dpt(1) = spt(1) - DimExtend

Set dimline1 = AddLine(dpt, upt) '左標線
dimline1.rotate spt, math.deg2rad(90 - fi)

upt(0) = ept(0): upt(1) = ept(1) + DimExtend
dpt(0) = ept(0): dpt(1) = ept(1) - DimExtend

Set dimline2 = AddLine(dpt, upt) '右標線
dimline2.rotate ept, math.deg2rad(90 - fi)

mpt(0) = mpt(0) + Dimtxtheight * 2 * math.degcos(fi - 90) + dx
mpt(1) = mpt(1) + Dimtxtheight * 2 * math.degsin(fi - 90)

Select Case DimScale

Case "m": Calc = 1000
Case "cm": Calc = 10

End Select

If DimContent = "" Then DimContent = math.getLength(spt, ept) / Calc '標註內容

Set txtobj = AddMixText(DimContent, mpt, Dimtxtheight, 2)
'txtobj.rotate mpt, Math.deg2rad(90 - fi)

If IsGroup = True Then

    Dim entColl As New Collection
    entColl.Add arrow1
    entColl.Add arrow2
    entColl.Add dimline1
    entColl.Add dimline2
    'entColl.Add txtobj
    Call BeGroup(entColl)

End If

End Sub

'****直接用內建FUNCTION****

Sub AddMLeader(ByVal spt, ByVal mpt, ByVal ept, ByVal prompt As String, ByVal txtheight As Double)

Dim txtpt(2) As Double

txtpt(0) = ept(0)
txtpt(1) = ept(1) + 0.7 * txtheight

Set arrow = PlotArrow(spt, mpt, math.getLength(spt, mpt) / 8)

Set txtobj = AddMixText(prompt, txtpt, txtheight, 1)
Call txtobj.GetBoundingBox(Min, Max)

ept(0) = Max(0)

Set lineObj = AddLine(mpt, ept)

Exit Sub

Dim entColl As New Collection

entColl.Add arrow
entColl.Add lineObj

Call BeGroup(entColl)

End Sub

'****寫的真夠亂的*****

Sub DrawTitle(ByVal lstUpPrompt, ByVal lstDownPrompt, ByVal CenterPoint, ByVal txtheight As Double, ByVal bottomline As Integer)

Dim txtpt(2) As Double

lstUpPrompt = Split(lstUpPrompt, ",")

For j = 0 To UBound(lstUpPrompt)

    txtpt(0) = CenterPoint(0)
    txtpt(1) = CenterPoint(1) + 1.5 * txtheight * (j + 1)
    strPrompt = lstUpPrompt(j)

    If j = 0 Then
        Set txtobj = AddMixText(strPrompt, txtpt, txtheight, 2, bottomline)
    Else
        Set txtobj = AddMixText(strPrompt, txtpt, txtheight, 2)
    End If
    
Next

If lstDownPrompt = "" Then Exit Sub

lstDownPrompt = Split(lstDownPrompt, ",")

For j = 0 To UBound(lstDownPrompt)

    txtpt(0) = CenterPoint(0)
    
    If i = 1 Then CenterPoint(1) = CenterPoint(1) + 0.5 * txtheight
    
    txtpt(1) = CenterPoint(1) - 1.5 * txtheight * (j + 1)
    
    Set txtobj = AddMixText(lstDownPrompt(j), txtpt, txtheight, 2)

Next

End Sub

'***可以用BLOCK取代****

Sub DrawWaterLevel(ByVal retStart, ByVal retEnd)

Dim spt(2) As Double
Dim ept(2) As Double
Dim vertices(4 * 3 - 1) As Double

X1 = retStart(0): Y1 = retStart(1)
X2 = retEnd(0): Y2 = retEnd(1)

Set lineObj = AddLine(retStart, retEnd)

L = Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2))

r = L / 10

Xm = (X1 + X2) / 2: Ym = (Y1 + Y2) / 2

vertices(0) = Xm: vertices(1) = Ym
vertices(3) = Xm - L / 20: vertices(4) = Ym + L / 10
vertices(6) = Xm + L / 20: vertices(7) = Ym + L / 10
vertices(9) = Xm: vertices(10) = Ym

Set plineobj = AddPolyLine(vertices)

For i = 1 To 4

    spt(0) = Xm - r / 5 * (5 - i) * (5 - i): spt(1) = Ym - L / 20 * i
    ept(0) = Xm + r / 5 * (5 - i) * (5 - i): ept(1) = Ym - L / 20 * i
    
    Set lineObj = AddLine(spt, ept)
    
Next

End Sub

Function InsertBlock(ByVal pt_insert, ByVal block_name As String, ByVal myscale As Double, Optional ByVal attr_coll) '20201112 add

Dim pt_origin(2) As Double

pt_origin(0) = 0: pt_origin(1) = 0: pt_origin(2) = 0 '圖塊插入基準點

Set myBlock = acadDoc.Blocks(block_name)
myBlock.Origin = tranPoint(pt_origin)

Set InsertBlock = mo.InsertBlock(tranPoint(pt_insert), block_name, myscale, myscale, myscale, 0)

If IsError(attr_coll) Then Exit Function

If TypeName(attr_coll) = "Variant()" Then Set attr_coll = myFunc.tranArray2Coll(attr_coll)

With InsertBlock

    If CADVer = "ICAD" Then
        Set myAttr = .GetAttributes
    Else
        myAttr = .GetAttributes
    End If
    
    For i = 1 To attr_coll.Count
        
        If attr_coll(i) <> "" Then myAttr(i - 1).TextString = attr_coll(i)
    
    Next
    
End With

End Function

Function getBlockAttr(ByVal BLK As Object) As Collection

Dim coll As New Collection

If CADVer = "ICAD" Then
    Set myAttr = BLK.GetAttributes
Else
    myAttr = BLK.GetAttributes
End If

For Each it In myAttr
    Debug.Print it.TagString
    coll.Add it.TextString, it.TagString
Next

Set getBlockAttr = coll

End Function

Function IntersectWith(ByVal PL1, ByVal PL2, ByVal mode As Byte, ByRef IsIntersect As Boolean)

IsIntersect = False

Dim coll As New Collection

If CADVer = "ICAD" Then
    
    On Error GoTo ERRORHANDLE '20210901 錯誤偵測
    Set retpt = PL1.IntersectWith(PL2, mode) 'acExtendThisEntity)
    
    For Each it In retpt
    
        coll.Add it.X
        coll.Add it.Y
        coll.Add it.Z
    
    Next
    
    If coll Is Nothing Then
ERRORHANDLE:
        'intersectWith = "EMPTY"
    Else
        IntersectWith = myFunc.tranColl2Array(coll)
        IsIntersect = True
    End If
    
Else

    retpt = PL1.IntersectWith(PL2, mode) 'acExtendThisEntity)
    
    If UBound(retpt) = -1 Then
       ' IntersectWith = "EMPTY"
    Else
        IntersectWith = retpt
        IsIntersect = True
    End If

End If


End Function

Sub setLayer(ByVal obj, ByVal layername As String)

acadDoc.Layers.Add layername

obj.Layer = layername

End Sub

