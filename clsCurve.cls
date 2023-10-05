VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Curve As Object
Private CL As Object
Private Radius As Double
Private i As Double
Private ls As Double
Private t As Double

Private Xip As Double
Private Yip As Double
Private Xbc As Double
Private Ybc As Double
Private Xec As Double
Private Yec As Double
Private Xc As Double
Private Yc As Double

Public txtheight As Double

Private collCurvePt As New Collection

Private CAD As New clsACAD
Private math As New clsMath
Private CLobj As New clsCL

Sub GetCurve()

CAD.acadDoc.Utility.prompt "請選擇單曲線!"

Set sset = CAD.CreateSSET(, "8", "Curve")

Set Curve = sset.Item(0)

Debug.Print TypeName(Curve) '.objecttype

Curve.Layer = "Curve"

Radius = Curve.Radius
ls = Curve.arclength

spt = Curve.startpoint
Xbc = spt(0): Ybc = spt(1)

ept = Curve.endpoint
Xec = ept(0): Yec = ept(1)

cpt = Curve.Center
Xc = cpt(0): Yc = cpt(1)

End Sub

Sub GetIP()

pt = CAD.GetPoint("請選擇單曲線的IP點")

Xip = pt(0): Yip = pt(1)

thida = math.deg2rad(90 - math.getAz(Xbc, Ybc, Xip, Yip))
thida2 = math.deg2rad(90 - math.getAz(Xip, Yip, Xec, Yec))

i = thida2 - thida

t = Abs(Radius * Tan(i / 2))

End Sub

Sub CreatePoint()

Dim mypt(2) As Double

fi = math.getAz(Xec, Yec, Xc, Yc)

mypt(0) = Xec
mypt(1) = Yec

CAD.acadDoc.Utility.prompt "建議放樣間距為20M=" & Int(ls / 20) & "個點位"
step = CAD.GetString("請輸入要擬合的單曲線點位個數:")

For j = 0 To step

    ang = Curve.totalangle / 2 / step * j
    
    dx = 2 * Radius * Sin(ang) * Sin(ang)
    dy = 2 * Radius * Sin(ang) * Cos(ang)
    
    Set ptObj = CAD.AddPointCO(Xec + dx, Yec + dy)
    
    ptObj.rotate mypt, math.deg2rad(-fi + 90)
    
    collCurvePt.Add ptObj

Next

If IsReverse(collCurvePt) Then Set collCurvePt = ReverseColl(collCurvePt)

End Sub

Function IsReverse(ByVal coll As Collection) As Boolean

If CL Is Nothing Then Exit Function

pts = CL.coordinates

co = 3
If TypeName(CL) Like "*LWPolyline" Then co = 2

X1 = pts(0)
Xn = pts(UBound(pts) - co + 1)

tmpcoor = coll(1).coordinates

Xpt1 = tmpcoor(0)

tmpcoor = coll(coll.Count).coordinates

Xptn = tmpcoor(0)

If X1 < Xn And Xpt1 < Xptn Then

    IsReverse = False

Else

    IsReverse = True

End If


End Function

Function ReverseColl(ByVal coll As Collection)

Dim newcoll As New Collection

For j = coll.Count To 1 Step -1

newcoll.Add coll(j)

Next

Set ReverseColl = newcoll

End Function

Sub GetCL()

CAD.acadDoc.Utility.prompt "請選擇中心線!"

Set sset = CAD.CreateSSET("CL")

Set CL = sset(0)

End Sub

Sub ChangeCLpoint()

Dim vertices() As Double

co = 3
If TypeName(CL) Like "*LWPolyline" Then co = 2

pts = CL.coordinates

ReDim vertices(((UBound(pts) + 1) / co - 1) * 3 + collCurvePt.Count * 3 - 1)

For j = 0 To UBound(pts) Step co

    If pts(j) = Xip And pts(j + 1) = Yip Then
    
    For Each pt In collCurvePt
    
        mycoordinates = pt.coordinates
        
        vertices(k) = mycoordinates(0)
        vertices(k + 1) = mycoordinates(1)
        
        k = k + 3
        
        pt.Delete
    
    Next
    
    Else
    
    vertices(k) = pts(j)
    vertices(k + 1) = pts(j + 1)
    
    k = k + 3
    
    End If

Next

Set newCL = CAD.AddPolyLine(vertices)

End Sub

Sub ClearAll()

Dim vertices(3 * 3 - 1) As Double

vertices(0) = Xbc
vertices(1) = Ybc
vertices(3) = Xip
vertices(4) = Yip
vertices(6) = Xec
vertices(7) = Yec

'Set oldCL = CAD.AddPolyLine(vertices)

If CAD.GetString("是否要刪除引用中心線(Y/N)") = "Y" Then CL.Delete

'Curve.Delete

End Sub

Sub DimCurve()

Dim mypt(2) As Double

mypt(0) = Xip
mypt(1) = Yip

leapt = CAD.GetPoint("Select the note point")

Set circleobj = CAD.AddCircle(mypt, txtheight / 5)

deltaI = math.rad2deg(i)

arr = Array("IP", "X=" & Round(Xip, 3), "Y=" & Round(Yip, 3), "R=" & Radius, "I=" & math.Convert_Degree(deltaI), "T=" & Round(t, 3), "Ls=" & Round(ls, 3))

For j = LBound(arr) To UBound(arr)

leapt(1) = leapt(1) - txtheight * 1.2

Set txtobj = CAD.AddMixText(arr(j), leapt, txtheight)

Next

End Sub

Sub GetPointFromCurve()

With Sheets("曲線")

    Lc = .Cells(2, .Columns.Count).End(xlToLeft).Column + 1
    
    If Lc = 2 Then Lc = 1
    
    .Cells(2, Lc) = "X"
    .Cells(2, Lc + 1) = "Y"
    
    r = 3
    
For Each pt In collCurvePt

    tmpcoor = pt.coordinates
    
    X = tmpcoor(0)
    Y = tmpcoor(1)

    .Cells(r, Lc) = Round(X, 3)
    .Cells(r, Lc + 1) = Round(Y, 3)
    
    r = r + 1

Next

End With

End Sub

Sub DeleteCurvePoint() 'debug

Dim vertices() As Double

pts = CL.coordinates

co = 3
If TypeName(CL) Like "*LWPolyline" Then co = 2

For j = 0 To UBound(pts) Step co

    Xpt = pts(j)
    Ypt = pts(j + 1)

    For Each CurvePt In collCurvePt
    
        tmpcoor = CurvePt.coordinates
        
        If Xpt = tmpcoor(0) And Ypt = tmpcoor(1) Then k = k - 1
        
    Next

Next

ReDim vertices((UBound(pts) + 1) / 3 - k * 3 - 1)

For j = 0 To UBound(pts) Step co

    Xpt = pts(j)
    Ypt = pts(j + 1)
    
    vertices(m) = Xpt
    vertices(m + 1) = Ypt

    For Each CurvePt In collCurvePt
    
        tmpcoor = CurvePt.coordinates
        
        If Xpt = tmpcoor(0) And Ypt = tmpcoor(1) Then m = m - 1
        
    Next
    
    m = m + 1

Next

Stop


End Sub

Sub PlotAlignmentArc(Xs, Ys, Xm, Ym, Xe, Ye, Radius)

Dim t As Double

thida = math.deg2rad(90 - math.getAz(Xs, Ys, Xm, Ym))
thida2 = math.deg2rad(90 - math.getAz(Xm, Ym, Xe, Ye))

deltaI = thida2 - thida

t = Abs(Radius * Tan(deltaI / 2))

reversefi = thida + math.deg2rad(180)

Xbc = Xm + t * Cos(reversefi)
Ybc = Ym + t * Sin(reversefi)

'Call NoteAlignment(Xbc, Ybc, Math.getAz(Xs, Ys, Xm, Ym), "BC,0+000")

Xec = Xm + t * Cos(thida2)
Yec = Ym + t * Sin(thida2)

'Call NoteAlignment(Xec, Yec, Math.getAz(Xm, Ym, Xe, Ye), "EC,0+100")

Dim cpt(2) As Double

bE = thida2 + math.deg2rad(90)
fe = thida + math.deg2rad(90)

If deltaI < 0 Then

    bE = thida2 + math.deg2rad(90)
    fe = thida + math.deg2rad(90)
    
    Set ar = CAD.acadDoc.ModelSpace.AddArc(cpt, Radius, bE, fe)
    
    cpt(0) = Xec
    cpt(1) = Yec

Else

    bE = thida2 + math.deg2rad(270)
    fe = thida + math.deg2rad(270)
    
    Set ar = CAD.acadDoc.ModelSpace.AddArc(cpt, Radius, fe, bE)
    
    cpt(0) = Xbc
    cpt(1) = Ybc

End If

ar.Move ar.startpoint, cpt

ar.Layer = "Curve"

End Sub

Private Sub NoteAlignment(ByVal X, ByVal Y, ByVal fi, ByVal lstNote)

tmp = Split(lstNote, ",")

txtheight = 50
w = txtheight * (UBound(tmp) + 3)

Dim spt(2) As Double
Dim ept(2) As Double

spt(0) = X
spt(1) = Y

ept(0) = spt(0) + math.degcos(fi - 90) * w
ept(1) = spt(1) + math.degsin(fi - 90) * w

Set ret = CAD.AddLine(spt, ept)

For i = 0 To UBound(tmp)
    
    ept(0) = spt(0) + math.degcos(fi - 90) * (w - 1 - i * 1.1 * txtheight)
    ept(1) = spt(1) + math.degsin(fi - 90) * (w - 1 - i * 1.1 * txtheight)
    
    Dim txtpt(2) As Double
    
    txtpt(0) = ept(0) + math.degcos(fi - 90) * 5
    txtpt(1) = ept(1) + math.degsin(fi - 90) * 5
    
    Set ret = CAD.AddMixText(tmp(i), txtpt, txtheight)

    ret.rotate txtpt, math.deg2rad(-fi + 90)

Next

End Sub

Private Sub Class_Initialize()

Set lay = CAD.acadDoc.Layers.Add("Curve")

lay.Color = 6

End Sub
