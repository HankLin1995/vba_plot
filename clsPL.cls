VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private CAD As New clsACAD
Private Math As New clsMath
Private pt As New clsPt
Private myfunc As New clsMyfunction

Public PL As Object
Public PLname As String '不是那麼重要
Private pts As Variant
Private co As Integer
Private totalLen As Double

Private sht_PL As Object

Const NUM_DIGIT As Integer = 4

Function rebuildPL()

Dim coll As New Collection
Dim ptsObj As New clsPTs_Table

myCoordinates = PL.Coordinates

For i = 0 To UBound(myCoordinates) Step co

    x = myCoordinates(i)
    y = myCoordinates(i + 1)
    diff = 0.01
    
    Z = ptsObj.getNearlyPointHeight(x, y, diff, nearX, nearY)

    coll.Add nearX
    coll.Add nearY
    coll.Add Z

Next

newCoordinates = myfunc.tranColl2Array(coll)

Dim newPts() As Double

ReDim newPts(0 To UBound(newCoordinates))

For i = 0 To UBound(newCoordinates)

    newPts(i) = newCoordinates(i)

Next

Set newPL = CAD.AddPolyLine(newPts)

newPL.Layer = PLname

PL.Delete

End Function


Sub createPLByRow(ByVal r As Integer)

With sht_PL

    PLname = .Cells(r, 1)
    co = .Cells(r, 2)
    vertices_str = Split(.Cells(r, 3), ",")

    Dim vertices() As Double
    ReDim vertices(LBound(vertices_str) To UBound(vertices_str))
    
    For i = LBound(vertices_str) To UBound(vertices_str)
    
        vertices(i) = CDbl(vertices_str(i))
    
    Next

End With

If co = 2 Then
    Set PL = CAD.AddLWPolyLine(vertices)
Else
    Set PL = CAD.AddPolyLine(vertices)
End If

CAD.acadDoc.Layers.Add (PLname)

PL.Layer = PLname

End Sub

Sub getPropertiesByPL(bPL) 'get the PL object

Set PL = bPL
pts = CAD.tranIPoints(PL.Coordinates)
totalLen = PL.Length
co = 3: If TypeName(PL) Like "*LWPolyline" Then co = 2
PLname = PL.Layer

End Sub

Function joinStrFromVertices()   '2021224

For i = LBound(pts) To UBound(pts)

    a = a & "," & Round(pts(i), NUM_DIGIT)

Next

joinStrFromVertices = mid(a, 2)

End Function

Sub AppendData(Optional ByVal r As Double)

With sht_PL

    If r = 0 Then r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    
    .Cells(r, 1) = PLname
    '.Cells(r, 2) = TypeName(PL)
    .Cells(r, 2) = co
    .Cells(r, 3) = joinStrFromVertices

End With

End Sub

Sub interpolatePLCoorMix() '20210521 從聚合線生高程

'rely on getPLNodeBySelect

node1 = getPLNodeBySelect(CAD.GetPoint("StartPoint's Node=?"))
H1 = CAD.GetString("StartPoint's Height=?")
node2 = getPLNodeBySelect(CAD.GetPoint("EndPoint's Node=?"))
Hn = CAD.GetString("EndPoint's Height=?")

For j = (node1 - 1) * co To (node2 - 1) * co + 1 - co Step co '是否要抑制前後兩點?

    X1 = pts(j): Y1 = pts(j + 1)
    X2 = pts(j + co): Y2 = pts(j + 1 + co)
    
    partLen = partLen + Math.getLengthCO(X1, Y1, X2, Y2)

Next

'the first point

X1 = pts((node1 - 1) * co)
Y1 = pts((node1 - 1) * co + 1)
Z = H1

cnt = 1

Call pt.getPropertiesByUser("inter_" & cnt, X1, Y1, Z, "POINT_INT")
Call pt.CreatePoint(0)

For i = (node1 - 1) * co To (node2 - 1) * co + 1 - co Step co

    cnt = cnt + 1
    
    X1 = pts(i): Y1 = pts(i + 1)
    X2 = pts(i + co): Y2 = pts(i + 1 + co)
    
    dLen = Math.getLengthCO(X1, Y1, X2, Y2)
    s1 = s1 + dLen
    s2 = partLen - s1

    Z = (Hn * s1 + H1 * s2) / (s1 + s2)
    
    Call pt.getPropertiesByUser("inter_" & cnt, X2, Y2, Round(Z, 3), "POINT_INT")
    Call pt.CreatePoint(0)
    Call pt.AppendData

Next

End Sub

Private Function getPLNodeBySelect(ByVal pt) '20210521

For j = 0 To UBound(pts) Step co

    cnt = cnt + 1
    X1 = pts(j): Y1 = pts(j + 1)
    If X1 = pt(0) And Y1 = pt(1) Then: getPLNodeBySelect = cnt: Exit Function

Next

End Function

Sub interpolatePLCoor() '20210521

H1 = CAD.GetString("StartPoint's Height=?")
Hn = CAD.GetString("EndPoint's Height=?")

X1 = pts(0)
Y1 = pts(1)
Z = H1

cnt = 1

Call pt.getPropertiesByUser("inter_" & cnt, X1, Y1, Z, "POINT_INT")
Call pt.CreatePoint(0)
Call pt.AppendData

For i = 0 To UBound(pts) - co Step co

    cnt = cnt + 1
    
    X1 = pts(i): Y1 = pts(i + 1)
    X2 = pts(i + co): Y2 = pts(i + 1 + co)
    
    dLen = Math.getLengthCO(X1, Y1, X2, Y2)
    s1 = s1 + dLen
    s2 = totalLen - s1
    
    Z = (Hn * s1 + H1 * s2) / (s1 + s2)
    
    Call pt.getPropertiesByUser("inter_" & cnt, X2, Y2, Round(Z, 3), "POINT_INT")
    Call pt.CreatePoint(0)
    Call pt.AppendData
    
Next

End Sub

Function interpolation(ByVal crossPT, ByVal IsFromTable As Boolean, ByVal CollPointTable)

midX = crossPT(0)
midY = crossPT(1)

arr = pts

For i = 0 To UBound(pts) - co Step co


    'If i = 22 Then Stop
        '---------------------------------------------------------------

   ' beforeX = Format(Round(arr(i), 4), "0.000")
   ' beforeY = Format(Round(arr(i + 1), 4), "0.000")
    
    beforeX = Round(arr(i), 4)
    beforeY = Round(arr(i + 1), 4)
    
    If co = 3 Then beforeH = arr(i + 2) '線對應點資料擴充模組
    If IsFromTable = True Then beforeH = getZfromCollPointTable(beforeX & ":" & beforeY, CollPointTable)
    
    'Debug.Print beforeX & ">" & beforeY
    
    'afterX = Format(Round(arr(i + co), 4), "0.000")
    'afterY = Format(Round(arr(i + co + 1), 4), "0.000")
    
    afterX = Round(arr(i + co), 4)
    afterY = Round(arr(i + co + 1), 4)
    
    If co = 3 Then afterH = arr(i + co + 2) '線對應點資料擴充模組
    If IsFromTable = True Then afterH = getZfromCollPointTable(afterX & ":" & afterY, CollPointTable)
    
    '------------------------------------------------------------------
    
    If Math.IsMiddle(beforeX, beforeY, midX, midY, afterX, afterY) Then '20200701新增fun

        s1 = Sqr((midX - beforeX) ^ 2 + (midY - beforeY) ^ 2)
        s2 = Sqr((midX - afterX) ^ 2 + (midY - afterY) ^ 2)
        
        'Call CAD.AddPointCO(beforeX, beforeY)
        'Call CAD.AddPointCO(afterX, afterY)
        
        If afterH = 0 Or beforeH = 0 Then
            interpolation = 0
        Else
            interpolation = (s1 * afterH + s2 * beforeH) / (s1 + s2)
        End If
    
        Exit Function
    
    End If

Next

End Function

''-----若點資料不在兩點之間，用最邊邊的高程當他的值-----------
'
'For i = 0 To UBound(arr) - 3 Step 3
'
'    LeftX = arr(i): LeftY = arr(i + 1)
'    Lefth = arr(i + 2)
'
'    If Lefth <> 0 Then Exit For
'
'Next
'
'For i = UBound(arr) - 3 To 0 Step -3
'
'    RightX = arr(UBound(arr) - 2): RightY = arr(UBound(arr) - 1)
'    Righth = arr(UBound(arr))
'
'    If Righth <> 0 Then Exit For
'
'Next
'
'LeftL = (LeftX - MidX) ^ 2 + (LeftY - MidY) ^ 2
'RightL = (RightX - MidX) ^ 2 + (RightY - MidY) ^ 2
'
'If LeftL < RightL Then
'    interpolation = Lefth
'Else
'    interpolation = Righth
'End If
'
''----------------------------------------------------


Function getZfromCollPointTable(ByVal findkey As String, ByVal collPT_Table)

On Error GoTo ERRORHANDLE
getZfromCollPointTable = collPT_Table(findkey)

Exit Function

ERRORHANDLE:
getZfromCollPointTable = 0
Debug.Print findkey

End Function

Private Sub Class_Initialize()

Set sht_PL = Sheets("特徵線")

End Sub

Sub createPTbyPLfromTable() '20210802

Dim ptsObj As New clsPTs_Table
Dim pt As New clsPt

Set coll = ptsObj.getCollPointTable_TYLin 'FIX_POINT

For j = 0 To UBound(pts) Step co

    cnt = cnt + 1
    X1 = pts(j): Y1 = pts(j + 1)
    
    mykey = X1 & ":" & Y1
    
    Z = getZfromCollPointTable(mykey, coll)

    If Z <> 0 Then
    
        Call pt.getPropertiesByUser("0", X1, Y1, Z, "TY-LIN")
        Call pt.CreatePoint(0.5)
        Call pt.AppendData
    
    End If

Next

End Sub

Sub checkZFromTable(ByVal r As Double) '20210802

'r是指radius

Dim ptsObj As New clsPTs_Table
Dim pt As New clsPt

Set coll = ptsObj.getCollPointTable

For j = 0 To UBound(pts) Step co

    cnt = cnt + 1
    X1 = pts(j): Y1 = pts(j + 1)
    
    mykey = Round(X1, 4) & ":" & Round(Y1, 4)
    
    Z = getZfromCollPointTable(mykey, coll)

    If IsEmpty(Z) Then
    
        Debug.Print mykey
    
        Call getMinZ(X1, Y1, Z, d)

        Call pt.getPropertiesByUser("999", X1, Y1, Z, "ADD")
        Call pt.CreatePoint(0.5)

        If d > 1 Then

        Set circleobj = CAD.AddCircleCO(X1, Y1, r)
        circleobj.Layer = "0"

        PL.Color = 3

        End If
    
    End If

Next

End Sub

Sub getMinZ(ByVal X1 As Double, ByVal Y1 As Double, ByRef Z, ByRef d)

Dim resultArray() As Variant

With Sheets("總表")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    ReDim resultArray(1 To lr - 2 + 1, 1 To 2) ' 設定陣列的大小為計算結果的數量
    
    For r = 2 To lr
    
        X0 = .Cells(r, 2)
        Y0 = .Cells(r, 3)
    
        Z = .Cells(r, 4)
        d = ((X1 - X0) ^ 2 + (Y1 - Y0) ^ 2) ^ 0.5

        resultArray(r - 1, 1) = d ' 假設計算X值的函數為 計算X值()
        resultArray(r - 1, 2) = Z ' 假設計算Y值的函數為 計算Y值()
    
    Next

End With

Dim myFun As New clsMyfunction

Call myFun.MergeSort_array(resultArray, 1)

Z = resultArray(1, 2)
d = resultArray(1, 1)

End Sub

Function addPointByPL()   '20221227

Dim ptObj As New clsPt

For j = 0 To UBound(pts) Step co

    x = Round(pts(j), 3)
    y = Round(pts(j + 1), 3)

    Call ptObj.getPropertiesByUser("放樣", x, y, x, y)
    Call ptObj.CreatePoint(0.5)
    Call ptObj.AppendData

Next

End Function
