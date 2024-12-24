VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myloc As String
Private CA As Double
Private BA As Double
Private FA As Double

Private CAD As New clsACAD
Private Math As New clsMath
Private myfunc As New clsMyfunction
Private shtCL As Object
Private shtReport As Object
Private shtLsec As Object

Private Sub Class_Initialize()

Set shtCL = Sheets("中心線")
Set shtReport = Sheets("土石方報表")
Set shtLsec = Sheets("縱斷面繪圖")

End Sub

Sub Report_Main()

Dim FinalArray() As Variant
Dim Num As Integer

arr = Array(8, 64, 120, 176)
column_len = 48
j = 1

With shtReport

.Visible = True

'==========RESET==============

For i = 0 To UBound(arr)

    .Cells(arr(i) + 45, 4).Resize(, 2).UnMerge '20211126
    .Cells(arr(i) + 45, 7).Resize(, 2).UnMerge '20211126
    .Cells(arr(i) + 45, 10).Resize(, 2).UnMerge '20211126

    For c = 1 To 9
    
        .Cells(arr(i), c).Resize(column_len, 1).ClearContents
    
    Next
    
Next

'========GETDATA===============

Call inputdata(FinalArray(), Num) '取得基本資料

For i = 0 To UBound(arr)

    For r = 0 To 40 Step 2

        If FinalArray(j, 3) = 3 And IsSteped = False Then 'Get a blank

            j = j - 1
            
            IsSteped = True
            
        Else

            .Cells(arr(i) + r, 1) = FinalArray(j, 1)
            
            .Cells(arr(i) + r, 3) = FinalArray(j, 5)
            .Cells(arr(i) + r, 6) = FinalArray(j, 6)
            '.Cells(arr(i) + r, 9) = FinalArray(j, 7)
            
            IsSteped = False

        End If
        
        r_fir = arr(i) + r - 2
        r_sec = arr(i) + r

        FirLoc = Math.TranLoc(.Cells(r_fir, 1))
        SecLoc = Math.TranLoc(.Cells(r_sec, 1))
        
        If r_fir <> arr(i) - 2 Then

            If SecLoc <> 0 Then
            
                .Cells(r_sec - 1, 2) = SecLoc - FirLoc
                .Cells(r_sec - 1, 2).NumberFormatLocal = "0.00"
                .Cells(r_sec - 1, 4) = (.Cells(r_fir, 3) + .Cells(r_sec, 3)) / 2
                .Cells(r_sec - 1, 5) = "=B" & r_sec - 1 & "*D" & r_sec - 1
                .Cells(r_sec - 1, 7) = (.Cells(r_fir, 6) + .Cells(r_sec, 6)) / 2
                .Cells(r_sec - 1, 8) = "=B" & r_sec - 1 & "*G" & r_sec - 1
                .Cells(r_sec - 1, 10) = (.Cells(r_fir, 9) + .Cells(r_sec, 9)) / 2
                .Cells(r_sec - 1, 11) = "=B" & r_sec - 1 & "*J" & r_sec - 1
                
                CA = CA + .Cells(r_sec - 1, 5)
                BA = BA + .Cells(r_sec - 1, 8)
                RA = RA + .Cells(r_sec - 1, 11)
                L = L + .Cells(r_sec - 1, 2)
            
            End If

        End If

        j = j + 1
        
        If j > Num Then GoTo Last

    Next
    
    .Cells(arr(i) + r + 2, 1) = "小計"
    .Cells(arr(i) + r + 3, 2) = L & " M"
    .Cells(arr(i) + r + 3 - 2, 4) = "小計"
    .Cells(arr(i) + r + 3, 5 - 1) = CA & " M3"
    .Cells(arr(i) + r + 3 - 2, 7) = "小計"
    .Cells(arr(i) + r + 3, 8 - 1) = BA & " M3"
    .Cells(arr(i) + r + 3 - 2, 10) = "小計"
    .Cells(arr(i) + r + 3, 11 - 1) = RA & " M3"
    
    .Cells(arr(i) + r + 3, 4).Resize(2, 2).Merge '20211126
    .Cells(arr(i) + r + 3, 7).Resize(2, 2).Merge '20211126
    .Cells(arr(i) + r + 3, 10).Resize(2, 2).Merge '20211126
    
    If FinalArray(j, 3) <> 3 Then j = j - 1 '控制下一頁頁首

Next

Last:

For i = 0 To UBound(arr)

    If i = UBound(arr) Then Exit For

    If r_sec >= arr(i) And r_sec < arr(i + 1) Then Exit For
    
Next

For k = 0 To UBound(arr)

    .Cells(arr(k) - 4, 8) = ""

Next

For j = 0 To i

    .Cells(arr(j) - 4, 8) = "第" & j + 1 & "頁共" & i + 1 & "頁"

Next

.Cells(arr(i) + 42 + 2, 1) = "合計"
.Cells(arr(i) + 42 + 3, 2) = L & " M"
.Cells(arr(i) + 42 + 3 - 2, 4) = "合計"
.Cells(arr(i) + 42 + 3, 5 - 1) = CA & " M3"
.Cells(arr(i) + 42 + 3 - 2, 7) = "合計"
.Cells(arr(i) + 42 + 3, 8 - 1) = BA & " M3"
.Cells(arr(i) + 42 + 3 - 2, 10) = "合計"
.Cells(arr(i) + 42 + 3, 11 - 1) = RA & " M3"

.Cells(arr(i) + 42 + 3, 4).Resize(2, 2).Merge '20211126
.Cells(arr(i) + 42 + 3, 7).Resize(2, 2).Merge '20211126
.Cells(arr(i) + 42 + 3, 10).Resize(2, 2).Merge '20211126

.PageSetup.PrintArea = "A1:L" & arr(i) + 48

End With

Call ReportOutput

MsgBox "CA=" & CA & "m3" & vbNewLine & "BA=" & BA & "m3"

End Sub

Sub inputdata(ByRef FinalArray() As Variant, ByRef Num)

Dim BasicArray() As Variant

With shtCL

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

j = 1

ReDim BasicArray(1 To lr, 1 To 7) '有放比較多欄位

    For r = 3 To lr
        
        If .Cells(r, 1) = "" Then Exit For
    
        BasicArray(j, 1) = .Cells(r, 1)
        BasicArray(j, 2) = Math.TranLoc(.Cells(r, 1))
        BasicArray(j, 3) = .Cells(r, 1).Font.ColorIndex
        BasicArray(j, 4) = .Cells(r, 2 + 2)
        BasicArray(j, 5) = Round(.Cells(r, 3 + 2), 2)
        BasicArray(j, 6) = Round(.Cells(r, 4 + 2), 2)
        BasicArray(j, 7) = Round(.Cells(r, 5 + 2), 2)
        
        j = j + 1
        
    Next

    ReDim FinalArray(1 To j - 1, 1 To 7)
    
    For k = 1 To j - 1
        
        For m = 1 To 7
    
            FinalArray(k, m) = BasicArray(k, m)
    
        Next
    
    Next

    Num = k - 1

End With

End Sub

Private Sub ReportOutput()

Set wb = Workbooks.Add

shtReport.Copy wb.Sheets(1)

Application.DisplayAlerts = False

For Each mysht In wb.Sheets

ActiveWindow.DisplayZeros = False

i = i + 1

If i > 1 Then mysht.Delete

Next

Application.DisplayAlerts = True

shtReport.Visible = 0 'xlSheetHidden

End Sub

Sub DrawCABA_Main() '將橫斷面說明的CL、CA、BA刪除後重整

Set sset = CAD.CreateSSET("Title", "8", "TITLE")

For Each it In sset
        
    If TypeName(it) = "IAcadBlockReference" Then
    
        myAttr = it.GetAttributes
        HSecLoc = myAttr(0).TextString
        
        s = returnCABA(HSecLoc)
        
        tmp = Split(s, ",")
        
        myAttr(1).TextString = tmp(0)
        myAttr(2).TextString = "挖方=" & Round(tmp(1), 2) & " m2"
        myAttr(3).TextString = "填方=" & Round(tmp(2), 2) & " m2"
        'myAttr(4).TextString = "CLSM=" & Round(tmp(3), 2) & " m2"
        
    
    End If

ERRORHANDLE:

Next

End Sub

Private Function returnCABA(ByVal HSecLoc As String)

With shtCL

    Set rng = .Cells.Find(HSecLoc)
    
    r = rng.Row
    
    deltaH = .Cells(r, 2 + 2)
    CA = .Cells(r, 3 + 2)
    BA = .Cells(r, 4 + 2)
    RA = .Cells(r, 5 + 2)
    
    returnCABA = deltaH & "," & CA & "," & BA & "," & RA

End With

End Function

Sub CalcCABA() '20210720 update

Set ssetAreas = CAD.CreateSSET("HA", "0", "Hatch")

For Each ssetArea In ssetAreas

    'TODO:
    '1.getMidPoint
    '2.searchLoc
    '3.getAreaCol
    '4.keyinArea
    
    r = getRowFromBorder(ssetArea)
    c = getColFromLayerName(ssetArea)
    
    If r <> "" And c <> "" Then shtCL.Cells(r, c) = ssetArea.area / 1000000

Next

End Sub

Function getColFromLayerName(ByVal ha)

With shtCL

    CD = ha.Layer '橫斷面-挖方
    
    targetColName = Split(CD, "-")(1)

    Set rng = .Rows(2).Find(targetColName)
    
    getColFromLayerName = rng.Column

End With

End Function

Function getRowFromBorder(ByVal ha)

    Call CAD.GetBoundingBox(ha, MinX, MinY, MaxX, MaxY)
    midX = (MinX + MaxX) / 2
    midY = (MinY + MaxY) / 2

    With shtCL

        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
        For r = 3 To lr
        
            BorderPTs = Split(.Cells(r, 3), ",")
            
            If UBound(BorderPTs) = 3 Then
        
                Border_minX = CDbl(BorderPTs(0))
                Border_minY = CDbl(BorderPTs(1))
                Border_maxX = CDbl(BorderPTs(2))
                Border_maxY = CDbl(BorderPTs(3))
                
                If midX >= Border_minX And midX <= Border_maxX And midY >= Border_minY And midY <= Border_maxY Then
                    
                    getRowFromBorder = r
                    Exit For
                    
                End If
            
            End If
            
        Next

    End With

End Function

'1.sort X coordinate
'2.define border
'3.decide X-value

Function getDLArea(ByVal EG, ByVal DL)

Set PLobj1 = RevisePL(EG)

PL1_X = myfunc.tranColl2Array(getcoll(PLobj1, "X"))
PL1_Y = myfunc.tranColl2Array(getcoll(PLobj1, "Y"))

PLobj1.Delete

Set PLobj2 = RevisePL(DL)
DL.Delete

PL2_X = myfunc.tranColl2Array(getcoll(PLobj2, "X"))
PL2_Y = myfunc.tranColl2Array(getcoll(PLobj2, "Y"))

'PLobj2.Delete

Call defineBorder(PL1_X, PL2_X, BL, BR)

getDLArea = calculateArea_test(PL1_X, PL1_Y, PL2_X, PL2_Y, BL, BR)

End Function

Function RevisePL(ByVal PLobj) '由左至右排序

Dim coll As New Collection

co = 3
If TypeName(PLobj) Like "*LWPolyline" Then co = 2  'IIcadLightWeightPolyline

arr = CAD.tranIPoints(PLobj.Coordinates)

Debug.Print TypeName(PLobj)

X1 = arr(0)
Xn = arr(UBound(arr) - co + 1)

If X1 > Xn Then
    Set RevisePL = CAD.ReverseLine(ByVal PLobj)
Else
    Set RevisePL = PLobj
End If

For i = 0 To UBound(arr) - co Step co

    X1 = arr(i)
    X2 = arr(i + co)
    
    If X1 = X2 Then arr(i + co) = arr(i + co) + 0.001 '若為垂直則微調往右偏移

Next

Set RevisePL = CAD.AddLWPolyLine(arr)
RevisePL.Layer = PLobj.Layer

End Function

Function calculateArea_test(ByVal PL1_X, ByVal PL1_Y, ByVal PL2_X, ByVal PL2_Y, ByVal BL, ByVal BR)

X_sort = myfunc.BubbleSort_array(myfunc.combineArray(PL1_X, PL2_X))

For i = LBound(X_sort) To UBound(X_sort) - 1
    
    X1 = X_sort(i)
    X2 = X_sort(i + 1)
    
    dx = X2 - X1
    
    If X1 >= BL And X2 <= BR Then
        
        Y1f = getYbyX(X1, PL1_X, PL1_Y)
        Y2f = getYbyX(X1, PL2_X, PL2_Y)
        
        dYf = Y1f - Y2f
        
        Y1b = getYbyX(X2, PL1_X, PL1_Y)
        Y2b = getYbyX(X2, PL2_X, PL2_Y)
        
        dYb = Y1b - Y2b

        If dYf > 0 And dYb < 0 Then '與基準線交叉(正到負)
        
            X_intersect = getIntersectX(X1, X2, Y1f, Y1b, Y2f, Y2b)
            dA = dYf * (X_intersect - X1) / 2
            'Set lineObj = CAD.AddLineCO(X1, Y1f, X1, Y2f)
            
        ElseIf dYf < 0 And dYb > 0 Then '與基準線交叉(負到正)
        
            X_intersect = getIntersectX(X1, X2, Y1f, Y1b, Y2f, Y2b)
            dA = dYb * (X2 - X_intersect) / 2
            'Set lineObj = CAD.AddLineCO(X2, Y1b, X2, Y2b)
            
        ElseIf dYf >= 0 And dYb >= 0 Then '皆在基準線下方
        
            dA = (dYf + dYb) * dx / 2
            'Set lineObj = CAD.AddLineCO(X1, Y1f, X1, Y2f)
            'Set lineObj = CAD.AddLineCO(X2, Y1b, X2, Y2b)
        
        Else
            dA = 0
        
        End If
        
        'Debug.Print "dYf=" & dYf & ",dYb=" & dYb & ",dA=" & dA
        sA = sA + dA
        
    End If
    
Next

calculateArea_test = Round(sA / 1000000, 2)

End Function

Function getIntersectX(ByVal X1 As Double, ByVal X2 As Double, _
                        ByVal Y1f As Double, ByVal Y1b As Double, _
                        ByVal Y2f As Double, ByVal Y2b As Double)

dx = X2 - X1

y1_slope = (Y1b - Y1f) / dx
y2_slope = (Y2b - Y2f) / dx

slope_change = Abs(y1_slope - y2_slope)
dY_f = Abs(Y1f - Y2f)

getIntersectX = X1 + dY_f / slope_change

End Function

Function getYbyX(ByVal x As Double, ByVal PL_X, ByVal PL_Y)

For i = LBound(PL_X) To UBound(PL_X) - 1
    
    X1 = PL_X(i)
    X2 = PL_X(i + 1)
    Y1 = PL_Y(i)
    Y2 = PL_Y(i + 1)
    
    If X1 = X2 Then X2 = X2 + 0.001: Stop
    
    If x >= X1 And x <= X2 Then
        
        s1 = x - X1
        s2 = X2 - x
        getYbyX = (s1 * Y2 + s2 * Y1) / (s1 + s2)
        
        Exit For
      
    End If
    
Next

End Function

Sub defineBorder(ByVal P1, ByVal P2, ByRef Border_Left, ByRef Border_right)

'check left

If P1(0) > P2(0) Then
    Border_Left = P1(0)
Else
    Border_Left = P2(0)
End If

'check right

If P1(UBound(P1)) < P2(UBound(P2)) Then
    Border_right = P1(UBound(P1))
Else
    Border_right = P2(UBound(P2))
End If

End Sub

Function getcoll(ByVal PLobj As Object, ByVal mode As String)

Dim coll As New Collection

co = 3
If TypeName(PLobj) Like "*LWPolyline" Then co = 2

arr = CAD.tranIPoints(PLobj.Coordinates)

j = 1: If mode = "X" Then j = 0

For i = 0 To UBound(arr) Step co

    coll.Add arr(i + j)

Next

Set getcoll = coll

End Function

