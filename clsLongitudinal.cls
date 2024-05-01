VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLongitudinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ACAD As Variant
Private sht As Worksheet
Private MyACAD As New clsACAD
Private myMath As New clsMath
Private myFunc As New clsMyfunction
'Private MyACAD As New clsICAD

Private AllHeightData() As Variant
Private TableData() As Variant
Private ControlData() As Variant
Private NoteData() As Variant
Private LocArr As Variant
Private BaseHeightPoint(2) As Double
Private viewpoint(2) As Double
Private BarStepHeight As Double

Private TotalLength As Double
Private TableRow As Integer
Private Max As Double, Min As Double
Private NeedStep As Integer
Public IsShowHeightTxt As Boolean

Public Xscale As Double
Public Yscale As Double
Public txtheight As Double
Public Interval As Double
Public VHeight As Double
Public startInterval As Double
Public TitleWidth As Double
Public TableMaxHeight As Double 'In order to limit the YHeight
Public sc As Integer, Lc As Integer
Public LocType As String
Public NUM_DIGIT As Byte

Private sr As Integer, lr As Integer
Private rLoc As Integer, rPlan As Integer, rSE As Integer, rSlope As Integer, rNote As Integer, rdeltaH As Integer

Private Sub Class_Initialize()

NUM_DIGIT = 2

Set sht = Sheets("縱斷面繪圖")

With sht

    Set rng = .Columns("A").Find("樁號")
    sr = rng.Row
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    rLoc = sr

End With

myLayerArr = Array("TABLE", "TEXT", "BAR", "L-INFO")
myLayerColorArr = Array(3, 6, 6, 7)

For Each it In myLayerArr

    With MyACAD.acadDoc
    
        .Layers.Add (it)
        If MyACAD.CADVer = "AUTOCAD" Then .Layers(it).Color = myLayerColorArr(i)
        
    End With
    
    i = i + 1

Next

End Sub

Sub ReadData() 'sr lr 本次產生後為模組內公用變數(private)

With sht

    'sr = 2 '樁號一定要放在第二列嗎?
    'lr = .Cells(100, 1).End(xlUp).Row

    TotalLength = (myMath.TranLoc(.Cells(sr, Lc)) - myMath.TranLoc(.Cells(sr, sc))) * 1000 'mm2m
    
    Call GetImportantRow '重要欄位
    Call GetAllHeightData '高程內容
    Call GetLocArrInTableData '取得LocArr
    Call GetTableData '表格內容
    Call GetControlData '控制點
    Call GetNoteData '旗標
    
    
    TableRow = UBound(TableData)
    
    Call CompareHeight
   'Call CreateHeightLayer
    
End With

'IsShowHeightTxt = True

'If MsgBox("不要秀出高程數字?", vbYesNo) = vbYes Then IsShowHeightTxt = False

End Sub

Private Sub GetImportantRow() '取得重要標示列位

With sht

    'sr = 2
    'lr = .Cells(100, 1).End(xlUp).Row

    For r = sr To lr
    
        TargetName = .Cells(r, 1)
        
        Select Case TargetName
        
            Case "計畫高": rPlan = r
            Case "坡降": rSlope = r
            Case "設計起始點": rSE = r
            Case "說明": rNote = r
            Case "挖填方": rdeltaH = r
            
        End Select
    
    Next
    
    If rPlan = 0 Then rPlan = InputBox("請輸入計畫高的列")
    If rSlope = 0 Then rSlope = InputBox("請輸入坡度的列")
    If rSE = 0 Then rSE = InputBox("請輸入設計起始點的列")
    If rNote = 0 Then rNote = InputBox("請輸入說明的列")
    If rdeltaH = 0 Then rdeltaH = InputBox("請輸入挖填方的列")
    
End With

End Sub

Private Sub GetAllHeightData() '取得所有高程資料

Dim HeightData() As Variant

With sht

    For r = 1 To lr
        If .Cells(r, 1).Value Like "*高" Then i = i + 1
    Next
    
    ReDim AllHeightData(1 To i, 1)
    ReDim HeightData(1 To Lc - 1)
    
    j = 1 '初始化
    
    For r = 1 To lr
    
        If .Cells(r, 1) Like "*高" Then

            AllHeightData(j, 0) = .Cells(r, 1)
            
            k = 1 '初始化
            
            For c = sc To Lc

                    HeightData(k) = .Cells(r, c)
                    k = k + 1

            Next
        
            AllHeightData(j, 1) = HeightData
            
            j = j + 1
        End If
        
    Next

End With

End Sub

Private Sub GetTableData() '取得表格資料

Dim coll As New Collection

With sht

    For r = sr To lr
    
        If .Cells(r, 1) = "" Then Exit For
        i = i + 1
        
    Next
    
    ReDim TableData(1 To i, 1 To 2) As Variant

    j = 1 '初始化

    For r = sr To i + sr - 1

        TableData(j, 1) = .Cells(r, 1)

        For c = sc To Lc

            If r = rLoc Then
            
                If .Cells(r, c) = .Cells(r, c + 1) Then
                    mystr = "(上)"
                ElseIf .Cells(r, c - 1) = .Cells(r, c) Then
                    mystr = "(下)"
                Else
                    mystr = ""
                End If
            
                coll.Add getLocValue(.Cells(r, c).Value) & mystr
                'coll.Add getLocValue(.Cells(r, c))
            Else
                coll.Add .Cells(r, c).Value
            End If
        
        Next

        ContentArr = myFunc.tranColl2Array(coll, 1)

        TableData(j, 2) = ContentArr
        
        Set coll = New Collection
        
        j = j + 1

    Next
    
    'Call GetLocArrInTableData '取得LocArr

End With

End Sub

Private Function getLocValue(ByVal s As Double) 'ByVal rng As Object) '20220915

Select Case LocType

Case "A": loc_format = "0+000"
Case "B": loc_format = "0K+000"

End Select

'getLocValue = Format(rng.Value, loc_format)
getLocValue = Format(s, loc_format)

If s Like "*.*" Then getLocValue = Format(s, loc_format) & Split(s, ".")(1)

'beforeLoc = rng.Offset(0, -1)
'backLoc = rng.Offset(0, 1)
'nowLoc = rng
'
'If nowLoc = backLoc Then
'    myStr = "(上)"
'ElseIf beforeLoc = nowLoc Then
'    myStr = "(下)"
'Else
'    myStr = ""
'End If
'
'getLocValue = getLocValue & myStr

End Function

Private Sub GetLocArrInTableData()

Dim coll As New Collection

With sht

    For c = sc To Lc
    
        coll.Add .Cells(rLoc, c).Value
    
    Next

    LocArr = myFunc.tranColl2Array(coll, 1)

End With

End Sub

Private Sub GetControlData() '取得設計起始點的左右標

Dim collwhere As New Collection
Dim collwhat As New Collection
Dim collHeight As New Collection
Dim collSlope As New Collection

With sht

    For c = sc To Lc
        
        If .Cells(rSE, c) <> "" Then
        
            collwhere.Add .Cells(sr, c).Value
            collwhat.Add .Cells(rSE, c).Value
            collHeight.Add Round(.Cells(rPlan, c).Value, 3)
            collSlope.Add .Cells(rSlope, c).Value
            
        End If
        
    Next
        
    ReDim ControlData(1 To collwhat.Count, 1 To 4)
    
    For i = 1 To collwhat.Count
    
        ControlData(i, 1) = collwhere.Item(i)
        ControlData(i, 2) = collwhat.Item(i)
        ControlData(i, 3) = collHeight.Item(i)
        ControlData(i, 4) = collSlope.Item(i)
        
    Next

End With

End Sub

Private Sub GetNoteData()

Dim collwhere As New Collection
Dim collsay As New Collection

With sht
    
    For c = sc To Lc

        If .Cells(rNote, c) <> "" Then
        
            collwhere.Add .Cells(rLoc, c)
            collsay.Add .Cells(rNote, c)
        
        End If
    Next

On Error Resume Next

ReDim NoteData(1 To collwhere.Count, 1 To 2)

For i = 1 To collwhere.Count

    NoteData(i, 1) = collwhere.Item(i)
    NoteData(i, 2) = collsay.Item(i)

Next

End With

End Sub

Sub GetScale()

'MsgBox "請移動至AutoCAD點選縱斷面視圖生成位置之左下角點"

On Error GoTo ERRORHANDLE

ret2 = MyACAD.GetPoint("點選縱斷面視圖生成位置之左下角點")

viewpoint(0) = ret2(0)
viewpoint(1) = ret2(1)

'Xscale = InputBox("請問X軸比例為" & vbCrLf & "1:")
'Yscale = InputBox("請問Y軸比例為" & vbCrLf & "1:")

Exit Sub

ERRORHANDLE:

viewpoint(0) = 0 'ret2(0)
viewpoint(1) = 0 'ret2(1)

End Sub

Private Sub CreateHeightLayer()

For i = 1 To UBound(AllHeightData())

    heightname = AllHeightData(i, 0)
    
    Set lay = MyACAD.acadDoc.Layers.Add(heightname)
    
    If heightname Like "*計*" Then
    
        lay.Color = 3
    
    End If
    
    'IntelliCAD.ActiveDocument.layers.Add (HeightName)

Next

End Sub

Sub DrawOuter()

Dim Ldpt(2) As Double, rupt(2) As Double
X1 = viewpoint(0)
Y1 = viewpoint(1)

For r = 1 To TableRow

    Y1 = viewpoint(1) + (Interval + VHeight) * (r - 1)

    Ldpt(0) = X1
    Ldpt(1) = Y1
    
    rupt(0) = X1 + TotalLength / Xscale + 2 * startInterval
    rupt(1) = Y1 + VHeight
    
    Set o = MyACAD.PlotRec(Ldpt, rupt)
    
    Call MyACAD.setLayer(o, "TABLE")
    
    'o.Layer = "TABLE" '"縱斷面-區帶格"
    
    Ldpt(0) = X1 - Interval - TitleWidth
    rupt(0) = X1 - Interval
    
    Set o = MyACAD.PlotRec(Ldpt, rupt)
    
    Call MyACAD.setLayer(o, "TABLE")
    
    'o.Layer = "TABLE" ' "縱斷面-區帶格"

Next

Call DrawOuterXLine

End Sub

Private Sub DrawOuterXLine() 'ByRef viewpoint As Variant)

Dim spt(2) As Double, ept(2) As Double

X1 = viewpoint(0)
Y1 = viewpoint(1)

For r = 1 To TableRow - 1
    
    Distance = 0
    
    For i = 1 To UBound(LocArr)
    
        Distance = (LocArr(i) - LocArr(1)) * 1000 / Xscale
        
        spt(0) = viewpoint(0) + Distance + startInterval
        spt(1) = Y1 + (Interval + VHeight) * (r)
        
        ept(0) = spt(0)
        ept(1) = spt(1) - Interval
        
        Set lineObj = MyACAD.AddLine(spt, ept)
        lineObj.Layer = "TABLE" ' "縱斷面-區帶格"
        
        'Set xlineobj = MyACAD.AddXLine(spt, ept)
        'xlineobj.Layer = "XLINE"
    
    Next

Next

End Sub

Private Sub FillInTableTitle() '資料區帶的標頭

Dim txtpt(2) As Double

X1 = viewpoint(0)
Y1 = viewpoint(1)

For r = 1 To TableRow

    Y1 = viewpoint(1) + (Interval + VHeight) * (r - 1)

    TableTitle = TableData(r, 1)
    
    txtpt(0) = X1 - Interval - TitleWidth / 2
    txtpt(1) = Y1 + VHeight / 2

    Set txtobj = MyACAD.AddText(TableTitle, txtpt, 4 / 3 * txtheight, 2) ' MyACAD.addtext(TableTitle, txtpt, 4 / 3 * txtheight, 2)
    
    txtobj.Layer = "TEXT" ' "縱斷面-區帶文字"
    
    Call MyACAD.FitTextWidth(txtobj, TitleWidth)

Next

End Sub

Sub FillInTable() '資料區帶的資料 '樁號需做特殊處理

Dim txtpt(2) As Double

Call FillInTableTitle

X1 = viewpoint(0) + startInterval
Y1 = viewpoint(1)

For r = 1 To TableRow - 1

    Y1 = viewpoint(1) + (Interval + VHeight) * (r - 1)

    ContentData = TableData(r, 2)
    'ncnt = 0
    For i = 1 To UBound(LocArr)
        
        Distance = (LocArr(i) - LocArr(1)) * 1000 / Xscale
        
        If i > 1 Then NextDistance = (LocArr(i) - LocArr(i - 1)) * 1000 / Xscale
        
        'Debug.Print NextDistance
        
        txtpt(0) = X1 + Distance
        txtpt(1) = Y1 + VHeight / 2
        
        s = ContentData(i)
        
        '******文字之間的特殊處理********

        storeMove = storeMove + NextDistance
        
        If i > 1 And storeMove < txtheight Then
            cnt = cnt + 1
            txtpt(0) = txtpt(0) + txtheight * cnt

        Else
            cnt = 0
            
        End If
        
        'Call MyACAD.AddPointCO(txtpt(0), txtpt(1))
        
'        If i > 1 And xmax < txtpt(0) Then txtpt(0) = txtpt(0) + txtheight
'        If IsMoved = True Then NextDistance = NextDistance - txtheight
'
'        If i > 1 And NextDistance < txtheight Then
'            txtpt(0) = txtpt(0) + txtheight
'            IsMoved = True
'        Else
'            IsMoved = False
'
'        End If
'
        '******繪製***************

        If IsNumeric(s) Then s = Round(s, NUM_DIGIT)
        
        Set txtobj = MyACAD.AddText(s, txtpt, txtheight, 2) ' MyACAD.addtext(s, txtpt, txtheight, 2)
        
        txtobj.rotate MyACAD.tranPoint(txtpt), 3.14 / 2
        
        txtobj.Layer = "TEXT" ' "縱斷面-區帶文字"
        
        'Call MyACAD.GetBoundingBox(txtobj, xmin, ymin, xmax, ymax)

        
    Next

Next

End Sub

Private Sub CompareHeight()

Max = 0
Min = 1000

For i = 1 To UBound(AllHeightData())

    heightArr = AllHeightData(i, 1)

    For j = 1 To UBound(heightArr)

        If heightArr(j) > Max Then Max = heightArr(j)
        If heightArr(j) < Min And heightArr(j) <> 0 Then Min = heightArr(j)
        
    Next

Next

Max = Int(Max) + 1
Min = Int(Min) - 1

End Sub

Sub DrawHeightBar() 'ByRef viewpoint As Variant)

Dim spt(2) As Double, ept(2) As Double, txtpt(2) As Double

BaseHeightPoint(0) = viewpoint(0) - Interval
BaseHeightPoint(1) = viewpoint(1) + (Interval + VHeight) * (TableRow)

'BarHeight = Int(TableMaxHeight - ((Interval + VHeight) * (TableRow + 1)))
BarStepHeight = 10 'mm

NeedStep = Int(((Max - Min) * 1000 / Yscale) / BarStepHeight)

'這裡太雞肋

'If NeedStep < 6 Then
    
'    msg = MsgBox("請再給予一個大一點的比例尺" & vbCrLf & "目前比例尺為： 1:" & Yscale & "有點太小了!!只有" & NeedStep & "階有點少", vbYesNo)
'    If msg = vbYes Then Yscale = InputBox("請問Y軸比例為" & vbCrLf & "1:")

'ElseIf NeedStep > 12 Then

'    msg = MsgBox("請再給予一個小一點的比例尺" & vbCrLf & "目前比例尺為： 1:" & Yscale & "有點太大了!!共有" & NeedStep & "階有點多", vbYesNo)
'    If msg = vbYes Then Yscale = InputBox("請問Y軸比例為" & vbCrLf & "1:")

'End If

NeedStep = Int(((Max - Min) * 1000 / Yscale) / BarStepHeight)

With MyACAD 'controlobject

    For i = 1 To NeedStep
    
        spt(0) = BaseHeightPoint(0)
        spt(1) = BaseHeightPoint(1) + i * BarStepHeight
        
        ept(0) = spt(0) - 15
        ept(1) = spt(1)
        
        Set lineObj = .AddLine(spt, ept)
    
        lineObj.Layer = "BAR" ' "縱斷面-高程條"
        
        'Set xlineobj = .AddXLine(spt, ept)
        'xlineobj.Layer = "XLINE"
    
        txtpt(0) = BaseHeightPoint(0) - 20
        txtpt(1) = BaseHeightPoint(1) + i * BarStepHeight
    
        s = Min + i * Yscale / (1000 / BarStepHeight)
    
        Set txtobj = MyACAD.AddText(s, txtpt, txtheight, 2)
        
        Call MyACAD.setLayer(txtobj, "BAR")
        'txtobj.Layer = "BAR" ' "縱斷面-高程條"
        
        
    Next
    
    ept(0) = spt(0): ept(1) = spt(1)
    spt(0) = spt(0): spt(1) = BaseHeightPoint(1) + BarStepHeight
    
    Set lineObj = .AddLine(spt, ept)
    
    lineObj.Layer = "BAR" ' "縱斷面-高程條"
    
    For i = 1 To NeedStep - 1
    
        spt(0) = BaseHeightPoint(0)
        spt(1) = BaseHeightPoint(1) + (i + 0.5) * BarStepHeight
        
        ept(0) = spt(0) - 6
        ept(1) = spt(1)
        
        Set lineObj = .AddLine(spt, ept)
        lineObj.Layer = "BAR" ' "縱斷面-高程條"
    
        txtpt(0) = BaseHeightPoint(0) - 12
        txtpt(1) = BaseHeightPoint(1) + i * 10 + 5
    
        s = Min + (i + 0.5) * Yscale / (1000 / BarStepHeight)
    
        Set txtobj = MyACAD.AddText(s, txtpt, txtheight, 2)
        txtobj.Layer = "BAR" ' "縱斷面-高程條"
    
    Next

End With

End Sub

Sub DrawHeight() '繪製POLYLINE

Dim vertices() As Double
Dim collLoc As New Collection
Dim collHeight As New Collection

X1 = viewpoint(0) + startInterval
Y1 = viewpoint(1)

With MyACAD.acadDoc

    For i = 1 To UBound(AllHeightData)
    
        heightname = AllHeightData(i, 0)
        HeightData = AllHeightData(i, 1)
        '-----20210706add-----
        
        .Layers.Add ("縱斷面-" & heightname)
        
        '---------------------
        
        .ActiveLayer = .Layers("縱斷面-" & heightname)
        
        For j = 1 To UBound(LocArr) '這裡有bug，可以修改為偵測SE值
        
            If HeightData(j) <> 0 Then
                
                IsDrawStarted = True
                
                collLoc.Add LocArr(j)
                collHeight.Add HeightData(j)
            
            ElseIf HeightData(j) = 0 And IsDrawStarted = True Then
                
                Call DrawHeightNow(collLoc, collHeight)
                IsDrawStarted = False
            
            End If
        
        Next
        
        If IsDrawStarted = True Then
        
            Call DrawHeightNow(collLoc, collHeight)
            IsDrawStarted = False
        
        End If
    
        '.ActiveLayer = .layers("0")
    
    Next

End With

End Sub

Private Sub DrawHeightNow(ByRef collLoc As Collection, ByRef collHeight As Collection)

Dim vertices() As Double, txtpt(2) As Double

ReDim vertices(3 * collLoc.Count - 1)

k = 0

showList = WorksheetFunction.RandBetween(2, collLoc.Count)

For j = 1 To collLoc.Count

    Distance = (collLoc.Item(j) - LocArr(1)) * 1000 / Xscale
    
    Xh = viewpoint(0) + Distance + startInterval
    Yh = BaseHeightPoint(1) + (collHeight.Item(j) - Min) * BarStepHeight * 100 / Yscale
    
    If IsShowHeightTxt = True Then
    
        txtpt(0) = Xh: txtpt(1) = Yh + 0.4
        Set txtobj = MyACAD.AddText(collHeight.Item(j), txtpt, 0.4 * txtheight, 2)

        txtobj.Layer = "TEXT"

    End If

    vertices(k + 3 * (j - 1)) = Xh
    vertices(k + 1 + 3 * (j - 1)) = Yh
    
    If j = showList Then
    
        Dim spt(2) As Double
        Dim mpt(2) As Double
        Dim ept(2) As Double
    
        spt(0) = Xh
        spt(1) = Yh
        mpt(0) = Xh + 5
        mpt(1) = Yh + 5
        ept(0) = Xh + 5
        ept(1) = Yh + 5
        
        prompt = MyACAD.acadDoc.ActiveLayer.Name
    
        Call MyACAD.AddMLeader(spt, mpt, ept, prompt, txtheight)
    
    End If

Next

Set plineobj = MyACAD.AddPolyLine(vertices)


Set collLoc = New Collection
Set collHeight = New Collection

End Sub

Sub FillInSlopeAndSE()

Dim spt(2) As Double, ept(2) As Double, cpt(2) As Double, txtpt(2) As Double
Dim Ldpt(2) As Double, rupt(2) As Double
Dim Arc As Object

With MyACAD

'.acaddoc.ActiveLayer = .acaddoc.Layers("縱斷面-施工資訊")
.acadDoc.ActiveLayer = .acadDoc.Layers("L-INFO")

X1 = viewpoint(0) + startInterval
Xbar = X1
Y1 = viewpoint(1) + (TableRow - 1) * (Interval + VHeight)

For i = 1 To UBound(ControlData)

    ConLoc = myMath.TranLoc(ControlData(i, 1))
    If i <> UBound(ControlData) Then ConLocNext = myMath.TranLoc(ControlData(i + 1, 1))
    mode = ControlData(i, 2)
    H = ControlData(i, 3)
    If i <> UBound(ControlData) Then hNext = ControlData(i + 1, 3)
    slope = ControlData(i, 4)
    
    Distance = (ConLoc - LocArr(1)) * 1000 / Xscale
    DistanceNext = (ConLocNext - ConLoc) * 1000 / Xscale
    
    spt(0) = X1 + Distance: spt(1) = Y1
    ept(0) = spt(0): ept(1) = spt(1) + VHeight
    cpt(0) = (spt(0) + ept(0)) / 2: cpt(1) = (spt(1) + ept(1)) / 2
    
    Set lineObj = .AddLine(spt, ept)
    
    Select Case mode
    
        Case "S"
            
            txtpt(0) = cpt(0) + VHeight / 5: txtpt(1) = cpt(1)
            Set txtobj = MyACAD.AddText(H, txtpt, txtheight, 2)
            Set Arc = .AddArc(cpt, VHeight / 2, -3.14 / 2, 3.14 / 2)
            
            spt(0) = cpt(0) + VHeight / 2: spt(1) = cpt(1)
            ept(0) = ept(0) + DistanceNext - VHeight / 2: ept(1) = cpt(1)
            Set lineObj = .AddLine(spt, ept)
            
        Case "C"
        
            lineObj.Delete
            txtpt(0) = cpt(0): txtpt(1) = cpt(1)
            Set txtobj = MyACAD.AddText(H, txtpt, txtheight, 2)
            Set circleobj = .AddCircle(cpt, VHeight / 2)
            
            spt(0) = cpt(0) + VHeight / 2: spt(1) = cpt(1)
            ept(0) = ept(0) + DistanceNext - VHeight / 2: ept(1) = cpt(1)
            Set lineObj = .AddLine(spt, ept)
            
        Case "E"
        
            txtpt(0) = cpt(0) - VHeight / 5: txtpt(1) = cpt(1)
            Set txtobj = MyACAD.AddText(H, txtpt, txtheight, 2)
            Set Arc = .AddArc(cpt, VHeight / 2, 3.14 / 2, -3.14 / 2)
    
    End Select

    txtobj.rotate .tranPoint(txtpt), 3.14 / 2
    
    Ldpt(0) = X1 + Distance
    Ldpt(1) = Y1 + VHeight + 2 * Interval
    rupt(0) = Ldpt(0) + DistanceNext
    rupt(1) = Ldpt(1) + 3 * Interval
    
    Xbar = Xbar + (ConLocNext - ConLoc) / Xscale * 1000
    
    Set Rec = MyACAD.PlotRec(Ldpt, rupt)
    
    If mode = "E" Then
    
    'Call MyACAD.Hatch(Rec, 1, "SOLID") 'BUG
    
    GoTo jump
     
    End If
     
    txtpt(0) = (spt(0) + ept(0)) / 2: txtpt(1) = (spt(1) + ept(1)) / 2 + VHeight / 4
    Set txtobj = MyACAD.AddText("L=" & (ConLocNext - ConLoc) & " M,Δh=" & Round((hNext - H), 3) & " M", txtpt, txtheight, 2)
    
    txtpt(0) = (spt(0) + ept(0)) / 2: txtpt(1) = (spt(1) + ept(1)) / 2 - VHeight / 4
    Set txtobj = MyACAD.AddText("S=1:" & slope, txtpt, txtheight, 2)
    
    txtpt(0) = (spt(0) + ept(0)) / 2: txtpt(1) = (spt(1) + ept(1)) / 2 + 0.5 * VHeight + 3.5 * Interval
    Set txtobj = MyACAD.AddText("施工區域", txtpt, 4 / 3 * txtheight, 2)
    
jump:
    
Next

End With

End Sub

Sub FillInNote() '縱斷面的說明標示(旗子)

Dim t As Integer
Dim vertices(3 * 3 - 1) As Double
Dim plineobj
Dim spt(2) As Double, ept(2) As Double, txtpt(2) As Double

X1 = viewpoint(0) + startInterval
Y1 = viewpoint(1)

Yn = TableRow * (Interval + VHeight) - Interval

On Error GoTo ERRORHANDLE

For i = 1 To UBound(NoteData)

    LocNote = myMath.TranLoc(NoteData(i, 1))
    
    Distance = (LocNote - LocArr(1)) * 1000 / Xscale
    
    Xn = X1 + Distance
    
    If i > 1 Then
        If LocNote - myMath.TranLoc(NoteData(i - 1, 1)) < 50 Then
            t = t + 1
        Else
            t = 0
        End If
    End If
    
    vertices(0) = Xn: vertices(1) = Y1 + Yn
    vertices(3) = Xn: vertices(4) = Y1 + Yn + NeedStep * 10 - t * (txtheight + Interval)
    vertices(6) = vertices(3) + 22 * Sqr(3): vertices(7) = vertices(4) + 22
    
    Set plineobj = MyACAD.AddPolyLine(vertices)
    
    plineobj.Layer = "L-INFO" '"縱斷面-施工資訊"
    
    NoteArr = Split(NoteData(i, 2), "、") 'NoteSeperate(NoteData(i, 2))
    
    For j = 0 To UBound(NoteArr)
    
        txtpt(0) = vertices(3) + Interval * (j - 1) + Interval
        txtpt(1) = vertices(4) + txtheight / 2 - (txtheight + Interval) * (j) + Interval
        
        s = "Sta:" & getLocValue(NoteData(i, 1)) & ":" & NoteArr(j)
        
        Set txtobj = MyACAD.AddText(s, txtpt, 4 / 3 * txtheight, 1)
        
        txtobj.rotate MyACAD.tranPoint(txtpt), 3.14 / 6
        
        txtobj.Layer = "L-INFO"  '"縱斷面-施工資訊"
        
        If j > 0 Then
        
            spt(0) = vertices(3): spt(1) = vertices(4) - (j) * (txtheight + Interval)
            ept(0) = vertices(6): ept(1) = vertices(7) - (j) * (txtheight + Interval)
            
            Set lineObj = MyACAD.AddLine(spt, ept)
            
            lineObj.Layer = "L-INFO" ' "縱斷面-施工資訊"
            
            t = t + 1
        
        End If
        
    Next
    
Next

ERRORHANDLE:

End Sub

Private Function FillInNoteSeperate(ByVal s As String) As Variant '說明標示分隔

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

FillInNoteSeperate = NoteArr

End Function

Sub TableIntroduce() '縱斷面圖比例標示

Dim txtpt(2)  As Double

'Dim attrcoll As New Collection

'txtpt(0) = viewpoint(0) - 100
'txtpt(1) = viewpoint(1)

'attrcoll.Add "V=1:" & Yscale
'attrcoll.Add "H=1:" & Xscale

'Set o = MyACAD.InsertBlock(txtpt, "LPF_SCALE", 1, attrcoll)

'o.Layer = "TITLE" '"比例尺"

'Exit Sub

'Dim txtpt(2)  As Double

txtpt(0) = viewpoint(0) - 100
txtpt(1) = viewpoint(1) + 50

Set txtobj = MyACAD.AddText("橫=1:" & Xscale, txtpt, 5 / 3 * txtheight, 1)

txtpt(1) = txtpt(1) - 20

Set txtobj = MyACAD.AddText("縱=1:" & Yscale, txtpt, 5 / 3 * txtheight, 1)

End Sub

Sub GetPlanDiff()

Dim collLoc As New Collection
Dim collPlanDiff As New Collection
Dim collBlank As New Collection
Dim TargetCD As String

Call GetImportantRow

With sht 'Sheets("縱斷面繪圖")

    For r = 2 To .Cells(2, 1).End(xlToRight).Column
    
        collLoc.Add Format(.Cells(2, r).Value, "0+000.0")
        'collPlanDiff.Add .Cells(rdeltaH, r).Value
        collPlanDiff.Add .Cells(rPlan, r).Value
    
    Next

End With

TargetCD = InputBox("請輸入渠道中心點CD碼名稱")
'TargetType = InputBox("請輸入渠道型式<A,B,C,D...擇一>")

With Sheets("橫斷面")

collBlank.Add 1

For r = 2 To .Cells(Rows.Count, 1).End(xlUp).Row

    If .Cells(r, 1) = "" Then collBlank.Add r

Next

collBlank.Add .Cells(Rows.Count, 1).End(xlUp).Row + 1

For i = 1 To collBlank.Count - 1

sr = collBlank.Item(i) + 1
er = collBlank.Item(i + 1) - 1

    For r = sr To er
    
        CD = .Cells(r, 3)
        myloc = .Cells(sr, 4)
        
        If CD = TargetCD Then
        
        For j = 1 To collPlanDiff.Count
        
            If collLoc.Item(j) = myloc Then
            
                .Cells(r, 4) = collPlanDiff.Item(j) & "," & TargetType(myMath.TranLoc(myloc))
                
                Exit For
            
            End If
        
        Next

        End If

    Next

Next

End With

End Sub

Function TargetType(ByVal loc As String)

With Sheets("渠道型式表")

For r = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row

loc_s = .Cells(r, 2)
loc_e = .Cells(r, 3)
loc_tmp = Val(loc)

mytype = .Cells(r, 5)

If loc_tmp >= loc_s And loc_tmp < loc_e Then

TargetType = mytype

End If

Next

End With

End Function

Sub ExportToCL_deltaH()

Call GetImportantRow

With sht

    mylc = .Cells(2, .Columns.Count).End(xlToLeft).Column
    
    For c = 2 To mylc
    
        tmpLoc = Format(.Cells(2, c), "0+000.0")
    
        tmpDiff = .Cells(rdeltaH, c)
        
        Set rng = Sheets("中心線").Cells.Find(tmpLoc)
        
        If Not rng Is Nothing Then Sheets("中心線").Cells(rng.Row, 4) = tmpDiff
        
    Next

End With

End Sub

