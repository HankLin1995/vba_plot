VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private CAD As New clsACAD
Private Math As New clsMath

Private arrLen() As Variant
Private arrLoc() As Variant
Private collCLPt As New Collection '與邊界之中心點
Private CL As Object '中心線
Private BL As Object '邊界
Private pts As Variant
Private co As Byte
Private totalLen As Double
Private shtCL As Object

Public PaperScale As Double '平面圖配置比例
Public WIDTH_COE As Double '文字距離取樣邊界係數
Public w As Double '半邊寬度
Public nowLoc As Double '起始樁號
Public IsLeftShow As Boolean '左標註
Public IsRightShow As Boolean '右標註
Public NeedReverse As Boolean '是否要逆轉
Public NeedBox As Boolean '文字外框
Public NeedDir As Boolean '方向旗標
Public BLnext As Double '流心連線距離

'Const WIDTH_COE = 1.5
Const NUM_DIGIT = 1
Const CL_LAYER_NAME = "CL"
Const CL_CROSS_LAYER_NAME = "CL_CROSS"
Const CL_CROSS_TRI_LAYER_NAME = "CL_CROSS_TRI"
Const CL_CROSS_TEXT_LAYER_NAME = "CL_CROSS_TXT"

Private Sub Class_Initialize()

Set shtCL = Sheets("中心線")

End Sub

Sub CrossLine_Main()

Dim nextL As Double
Dim storeLen As Double

CAD.acadDoc.Layers.Add CL_CROSS_LAYER_NAME
CAD.acadDoc.Layers.Add CL_CROSS_TEXT_LAYER_NAME

IsFirst = True

'Call getCenterLine(pts, CO, totalLen)
'Call getLoc(totalLen)

k = 1
nextL = arrLen(0)

For i = 0 To UBound(pts) - co Step co

    X0 = pts(i): Y0 = pts(i + 1)
    X1 = pts(i + co): Y1 = pts(i + 1 + co)

    ptLen = Sqr((X1 - X0) ^ 2 + (Y1 - Y0) ^ 2)  '兩點距

    fi = Math.getAz(X0, Y0, X1, Y1)
    
    If IsFirst Then
    
        myloc = nowLoc 'Format(nowLoc + 0, "0K+000")
        Call DrawCrossLine(X0, Y0, fi, myloc)
        IsFirst = False
    
    End If
    
    sumLen = sumLen + ptLen

    If sumLen <= nextL Then storeLen = storeLen + ptLen '不夠用先存起來

    Do Until sumLen <= nextL
        
        moveL = nextL
        
        If storeLen <> 0 Then moveL = nextL - storeLen: storeLen = 0

        sumLen = sumLen - nextL

        If k <= UBound(arrLen) Then nextL = arrLen(k) '下一單距
        
        dx = Math.degcos(fi) * (moveL)
        dy = Math.degsin(fi) * (moveL)
        
        Xnow = X0 + dx: Ynow = Y0 + dy
        X0 = Xnow: Y0 = Ynow
        On Error GoTo errhandle
        myloc = arrLoc(k - 1)
        
        Call DrawCrossLine(X0, Y0, fi, myloc)

        k = k + 1
    
    Loop

    storeLen = sumLen
    
Next

myloc = nowLoc + totalLen ' Format(nowLoc + totalLen, "0K+000")

Call DrawCrossLine(X1, Y1, fi, myloc)

errhandle:

End Sub

Sub CrossLine_Main_Output()

Dim nextL As Double
Dim storeLen As Double
Dim pt As New clsPt

CAD.acadDoc.Layers.Add CL_CROSS_LAYER_NAME

IsFirst = True

'Call getCenterLine(pts, CO, totalLen)
'Call getLoc(totalLen)

k = 1
nextL = arrLen(0)

For i = 0 To UBound(pts) - co Step co

    X0 = pts(i): Y0 = pts(i + 1)
    X1 = pts(i + co): Y1 = pts(i + 1 + co)

    ptLen = Sqr((X1 - X0) ^ 2 + (Y1 - Y0) ^ 2)  '兩點距

    fi = Math.getAz(X0, Y0, X1, Y1)
    
    If IsFirst Then
    
        myloc = nowLoc 'Format(nowLoc + 0, "0K+000")
        Call pt.getPropertiesByUser(X0, X0, Y0, Y0, Format(myloc, "0K+000"))
        Call pt.CreatePoint(0.5)
        Call pt.AppendData
        'Call DrawCrossLine(X0, Y0, fi, myloc)
        IsFirst = False
    
    End If
    
    sumLen = sumLen + ptLen

    If sumLen <= nextL Then storeLen = storeLen + ptLen '不夠用先存起來

    Do Until sumLen <= nextL
        
        moveL = nextL
        
        If storeLen <> 0 Then moveL = nextL - storeLen: storeLen = 0

        sumLen = sumLen - nextL

        If k <= UBound(arrLen) Then nextL = arrLen(k) '下一單距
        
        dx = Math.degcos(fi) * (moveL)
        dy = Math.degsin(fi) * (moveL)
        
        Xnow = X0 + dx: Ynow = Y0 + dy
        X0 = Xnow: Y0 = Ynow
        On Error GoTo errhandle
        myloc = arrLoc(k - 1)
        
        'Call DrawCrossLine(X0, Y0, fi, myloc)
        Call pt.getPropertiesByUser(X0, X0, Y0, Y0, Format(myloc, "0K+000"))
        Call pt.CreatePoint(0.5)
        Call pt.AppendData
        k = k + 1
    
    Loop

    storeLen = sumLen
    
Next

myloc = nowLoc + totalLen ' Format(nowLoc + totalLen, "0K+000")

'Call DrawCrossLine(X1, Y1, fi, myloc)
Call pt.getPropertiesByUser(X1, X1, Y1, Y1, Format(myloc, "0K+000"))
Call pt.CreatePoint(0.5)
Call pt.AppendData

errhandle:

End Sub

Sub getCenterLine() 'pts, CO, totalLen)

If CL Is Nothing Then
    
    Set sset = CAD.CreateSSET("CL")
    Set CL = sset(0)
    If NeedReverse Then Set CL = CAD.ReverseLine(CL)
    Call CAD.setLayer(CL, CL_LAYER_NAME)
    
'    CL.Layer = CL_LAYER_NAME '"平面圖-中心線" '"中心樁"
    
End If

If CAD.CADVer = "ICAD" Then
    pts = CAD.tranIPoints(CL.Coordinates)
Else
    pts = CL.Coordinates
End If

totalLen = CL.Length
co = 3
If TypeName(CL) Like "*LWPolyline" Then co = 2

End Sub

Sub DrawCrossLine(ByVal X0 As Double, ByVal Y0 As Double, ByVal fi As Double, ByVal myloc As String)

CAD.acadDoc.ActiveLayer = CAD.acadDoc.Layers(CL_CROSS_TEXT_LAYER_NAME) '("平面圖-橫斷樁")

Dim spt(2) As Double
Dim ept(2) As Double

spt(0) = X0 + Math.degcos(fi - 90) * w
spt(1) = Y0 + Math.degsin(fi - 90) * w

ept(0) = X0 + Math.degcos(fi + 90) * w
ept(1) = Y0 + Math.degsin(fi + 90) * w

Set lineObj = CAD.AddLine(spt, ept)

Call CAD.setLayer(lineObj, CL_CROSS_LAYER_NAME)

'Call CAD.setLayer(lineObj, "平面圖-橫斷樁L")

'lineObj.Layer = "平面圖-橫斷樁L" '20210901

If NeedDir Then '方向旗標

    Call DrawDirection(spt, fi, -2 / 1000 * PaperScale * 5 / 16, 2 / 1000 * PaperScale)
    Call DrawDirection(ept, fi, 2 / 1000 * PaperScale * 5 / 16, 2 / 1000 * PaperScale)
    
End If

Call CAD.SetXdataToObj(lineObj, Format(myloc, "0+000.0")) ' Replace(myloc, "K", ""))

If IsRightShow Then Call DrawCrossLoc(X0, Y0, myloc, fi, 90, 1)
If IsLeftShow Then Call DrawCrossLoc(X0, Y0, myloc, fi, -90, 3)

With shtCL

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    
    .Cells(lr, 1) = Replace(myloc, "K", "")
    .Cells(lr, 2) = spt(0)
    '.Cells(lr, 3) = spt(1)

End With

End Sub

Private Sub DrawDirection(dpt, fi, dx, dy)

Dim spt(2) As Double
Dim Tri As Object
Dim vertices(4 * 3 - 1) As Double

x = dpt(0): y = dpt(1)

vertices(0) = x: vertices(1) = y
vertices(3) = x: vertices(4) = y + dy 'w / 5
vertices(6) = x + dx: vertices(7) = y
vertices(9) = x: vertices(10) = y

Set Tri = CAD.AddPolyLine(vertices)
Tri.rotate CAD.tranPoint(dpt), Math.deg2rad(-fi)

Call CAD.setLayer(Tri, CL_CROSS_TRI_LAYER_NAME)

Call CAD.Hatch(Tri, 1, "SOLID")

End Sub

Private Sub DrawCrossLoc(X0, Y0, myloc, fi, ByVal rotangle As Double, ByVal alignmode As Byte)

Dim tept(2) As Double

tept(0) = Round(X0 + Math.degcos(fi + rotangle) * w * WIDTH_COE, 3)
tept(1) = Round(Y0 + Math.degsin(fi + rotangle) * w * WIDTH_COE, 3)

myloc = Format(myloc, "0K+000.0") '20210709

'Set txtobj = CAD.AddMixText(myloc, tept, w / 5, alignmode)
Set txtobj = CAD.AddMixText(myloc, tept, 4 * PaperScale / 1000, alignmode)
If NeedBox Then Set boxobj = CAD.AddTextBox(txtobj)

txtobj.rotate CAD.tranPoint(tept), Math.deg2rad(-fi)
If NeedBox Then boxobj.rotate CAD.tranPoint(tept), Math.deg2rad(-fi)

End Sub

Sub getLoc() 'ByVal totalLen As Double)

Dim sumLen As Double
Dim c As Integer
IsAsk = True

'=======ClearALL======

With shtCL

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
If lr > 3 Then .Range("A3:" & "G" & lr).ClearContents

End With

'====================

Do While sumLen < totalLen

    If IsAsk = True Then
    
        nextLen = CAD.GetString("起迄樁號為 " & Format(nowLoc + 0, "0+000") & "~" & Format(nowLoc + totalLen, "0+000") & vbNewLine & _
                             "請輸入下一樁的單距，目前為" & Format(nowLoc + sumLen, "0+000"))
    
    End If
    
    If nextLen Like "*-" Then
        nextLen = Val(LEFT(nextLen, Len(nextLen) - 1)): IsAsk = False
    Else
        nextLen = Val(nextLen)
    End If

    sumLen = sumLen + nextLen
    
    ReDim Preserve arrLoc(c)
    ReDim Preserve arrLen(c)
    
    arrLen(c) = nextLen
    arrLoc(c) = nowLoc + sumLen 'Format(nowLoc + sumLen, "0K+000")
    
    c = c + 1

Loop

End Sub

Sub getLocXLS() '最後一個樁號會被捨棄(預設為最尾巴)

Dim collmyloc As New Collection

Set sht = Sheets("中心線")

With sht

    lr = .Cells(Rows.Count, 1).End(xlUp).Row

    For r = 3 To lr - 1

        collmyloc.Add Math.TranLoc(.Cells(r, 1))

    Next

    If lr > 3 Then .Range("A3:" & "G" & lr).ClearContents

End With

nowLoc = collmyloc(1)

For i = 1 To collmyloc.Count - 1

    ReDim Preserve arrLoc(i - 1)
    ReDim Preserve arrLen(i - 1)

    Interval = collmyloc(i + 1) - collmyloc(i)
    sumInterval = sumInterval + Interval

    arrLen(i - 1) = Interval
    arrLoc(i - 1) = Format(nowLoc + sumInterval, "0K+000.0")

Next
    
End Sub

Sub BorderLine_Main()

Dim nextL As Double
Dim storeLen As Double

IsFirst = True

Call getBorderLine(pts, co, totalLen)

k = 1
nextL = BLnext

For i = 0 To UBound(pts) - co Step co

    X0 = pts(i): Y0 = pts(i + 1)
    X1 = pts(i + co): Y1 = pts(i + 1 + co)

    ptLen = Sqr((X1 - X0) ^ 2 + (Y1 - Y0) ^ 2)  '兩點距

    fi = Math.getAz(X0, Y0, X1, Y1)
    
    If IsFirst Then
    
        Call getCLpt(X0, Y0, fi): IsFirst = False
    
    End If
    
    sumLen = sumLen + ptLen

    If sumLen <= nextL Then storeLen = storeLen + ptLen '不夠用先存起來

    Do Until sumLen <= nextL
        
        moveL = nextL
        
        If storeLen <> 0 Then moveL = nextL - storeLen: storeLen = 0

        sumLen = sumLen - nextL

        'If k <= UBound(arrLen) Then nextL = arrLen(k) '下一單距
        
        dx = Math.degcos(fi) * (moveL)
        dy = Math.degsin(fi) * (moveL)
        
        Xnow = X0 + dx: Ynow = Y0 + dy
        X0 = Xnow: Y0 = Ynow
        
        'Call CAD.AddPointCO(X0, Y0)
        
        Call getCLpt(X0, Y0, fi)

        k = k + 1
    
    Loop

    storeLen = sumLen
    
Next

myloc = Format(nowLoc + totalLen, "0K+000")

Call getCLpt(X1, Y1, fi)

End Sub

Sub getCLpt(ByVal X As Double, ByVal Y As Double, ByVal fi As Double)

Dim spt(2) As Double
Dim ept(2) As Double
Dim midpt(2) As Double

spt(0) = X + Math.degcos(fi + 90) * 100
spt(1) = Y + Math.degsin(fi + 90) * 100

ept(0) = X + Math.degcos(fi - 90) * 100
ept(1) = Y + Math.degsin(fi - 90) * 100

Set CrossLine = CAD.AddLine(spt, ept)

On Error Resume Next
'
'retpt1 = CrossLine.IntersectWith(CL, 1) 'acExtendThisEntity)
'retpt2 = CrossLine.IntersectWith(BL, 1) ' acExtendThisEntity)

Dim IsIntersect1 As Boolean
Dim IsIntersect2 As Boolean

retpt1 = CAD.IntersectWith(CrossLine, CL, 1, IsIntersect1)
retpt2 = CAD.IntersectWith(CrossLine, BL, 1, IsIntersect2)

If IsIntersect1 And IsIntersect2 Then

midpt(0) = (retpt1(0) + retpt2(0)) / 2
midpt(1) = (retpt1(1) + retpt2(1)) / 2

If midpt(0) <> 0 Then collCLPt.Add midpt

CrossLine.Delete

End If

End Sub

Private Sub getBorderLine(pts, co, tl)

Set sset = CAD.CreateSSET("CL")
Set CL = sset(0)

pts = CAD.tranIPoints(CL.coordinates)
tl = CL.Length
co = 3
If TypeName(CL) Like "*LWPolyline" Then co = 2

Set sset = CAD.CreateSSET("BL")
Set BL = sset(0)

End Sub

Sub DrawCenterLine()

Dim vertices() As Double
ReDim vertices(3 * collCLPt.Count - 1)

For i = 1 To collCLPt.Count

    pt = collCLPt.Item(i)
    
    vertices(3 * (i - 1)) = pt(0)
    vertices(3 * (i - 1) + 1) = pt(1)

Next

Set CL = CAD.AddPolyLine(vertices)

Call CAD.setLayer(CL, "CL")

'CL.Layer = "CL" ' "平面圖-中心線"

End Sub

Sub ClearStation()

MsgBox "請移動至CAD框選要刪除的樁號圖層", vbInformation

Set mySSET = CAD.CreateSSET(, "8", CL_CROSS_LAYER_NAME)

For Each it In mySSET

    it.Delete

Next

CAD.acadDoc.Utility.prompt "已經清除完成!" & vbNewLine

With Sheets("中心線")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    .Range("A3:B" & lr).ClearContents

End With

End Sub

Sub getDataByRng()

tmp = Split(shtCL.Range("I1"), ",")
w = tmp(0)
nowLoc = tmp(1)
IsLeftShow = tmp(2)
IsRightShow = tmp(3)
NeedBox = tmp(4)
NeedDir = tmp(5)
PaperScale = tmp(6)
WIDTH_COE = tmp(7)

End Sub

Sub setDataByUser()

arr = Array(w, nowLoc, IsLeftShow, IsRightShow, NeedBox, NeedDir, PaperScale, WIDTH_COE)

shtCL.Range("I1") = Join(arr, ",")

End Sub

Sub getSpecificLoc(ByVal x As Double, ByVal y As Double)

tmp = Split(shtCL.Range("I1"), ",")
w = tmp(0)
nowLoc = tmp(1)
IsLeftShow = tmp(2)
IsRightShow = tmp(3)
NeedBox = tmp(4)
NeedDir = tmp(5)
PaperScale = tmp(6)
WIDTH_COE = tmp(7)

For i = 0 To UBound(pts) - co Step co

    X0 = pts(i): Y0 = pts(i + 1)
    X1 = pts(i + co): Y1 = pts(i + 1 + co)

    ptLen = Math.getLengthCO(X0, Y0, X1, Y1) '兩點距

    If Math.IsMiddle(X0, Y0, x, y, X1, Y1) Then
    
        fi = Math.getAz(X0, Y0, X1, Y1)
        ptLen = Math.getLengthCO(X0, Y0, x, y)
        sumLen = sumLen + ptLen
        myloc = Format(nowLoc + sumLen, "0K+000.0")
        
        Call DrawCrossLine(x, y, fi, myloc)
        
        Exit Sub

    Else
        
        sumLen = sumLen + ptLen
       
    End If

Next

End Sub
