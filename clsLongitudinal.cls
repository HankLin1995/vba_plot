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
Private mymath As New clsMath
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
Private IsShowHeightTxt As Boolean

Public Xscale As Double
Public Yscale As Double
Public txtheight As Double
Public Interval As Double
Public VHeight As Double
Public startInterval As Double
Public TitleWidth As Double
Public TableMaxHeight As Double 'In order to limit the YHeight
Public sc As Integer, Lc As Integer

Private sr As Integer, lr As Integer
Private rPlan As Integer, rSE As Integer, rSlope As Integer, rNote As Integer, rdeltaH

Private Sub Class_Initialize()

Set sht = Sheets("�a�_��ø��")

'Set ACAD = MyACAD.acaddoc  'AutoCAD.ActiveDocument

'Call MyACAD.Createfontstyle

End Sub

Sub ReadData() 'lr �������ͫᬰ�Ҳդ������ܼ�(private)

With sht

    sr = 2
    'lc = .Cells(sr, 1).End(xlToRight).Column
    lr = .Cells(100, 1).End(xlUp).Row

    TotalLength = (mymath.TranLoc(.Cells(sr, Lc)) - mymath.TranLoc(.Cells(sr, sc))) * 1000 'mm2m
    
    Call GetImportantRow '���n���
    Call GetAllHeightData '���{���e
    Call GetTableData '���椺�e
    Call GetControlData '�����I
    Call GetNoteData '�X��
    
    TableRow = UBound(TableData)
    
    Call CompareHeight
   'Call CreateHeightLayer
    
End With

IsShowHeightTxt = True

If MsgBox("���n�q�X���{�Ʀr?", vbYesNo) = vbYes Then IsShowHeightTxt = False

End Sub

Private Sub GetImportantRow() '���o���n�ХܦC��

With sht

    sr = 2
    lr = .Cells(100, 1).End(xlUp).Row

    For r = sr To lr
    
        TargetName = .Cells(r, 1)
        
        Select Case TargetName
        
            Case "�p�e��": rPlan = r
            Case "�Y��": rSlope = r
            Case "�]�p�_�l�I": rSE = r
            Case "����": rNote = r
            Case "�����": rdeltaH = r
            
        End Select
    
    Next
    
    If rPlan = 0 Then rPlan = InputBox("�п�J�p�e�����C")
    If rSlope = 0 Then rSlope = InputBox("�п�J�Y�ת��C")
    If rSE = 0 Then rSE = InputBox("�п�J�]�p�_�l�I���C")
    If rNote = 0 Then rNote = InputBox("�п�J�������C")
    If rdeltaH = 0 Then rdeltaH = InputBox("�п�J����誺�C")
    
End With

End Sub

Private Sub GetAllHeightData() '���o�Ҧ����{���

Dim HeightData() As Variant

With sht

    For r = 1 To lr
        If .Cells(r, 1).Value Like "*��" Then i = i + 1
    Next
    
    ReDim AllHeightData(1 To i, 1)
    ReDim HeightData(1 To Lc - 1)
    
    j = 1 '��l��
    
    For r = 1 To lr
    
        If .Cells(r, 1) Like "*��" Then
            'AllHeightData(j, 0) = r
            AllHeightData(j, 0) = .Cells(r, 1)
            
            k = 1 '��l��
            
            For c = sc To Lc
                
                'HeightData(k, 0) = .Cells(2, c) '��Z
                'HeightData(k, 1) = HeightData(k, 0) + .Cells(2, c) '�ֶZ
                
                'If Not .Cells(2, c) Like "*(*" Then
                    HeightData(k) = .Cells(r, c)
                    k = k + 1
                'End If
                
                'HeightData(k) = .Cells(r, c) '���{
                'k = k + 1
            Next
        
            AllHeightData(j, 1) = HeightData
            
            j = j + 1
        End If
        
    Next

End With

End Sub

Private Sub GetTableData() '���o������

Dim coll As New Collection

With sht

    For r = sr To lr
    
        If .Cells(r, 1) = "" Then Exit For
        i = i + 1
        
    Next
    
    ReDim TableData(1 To i, 1 To 2) As Variant

    j = 1 '��l��

    For r = sr To i + sr - 1

        TableData(j, 1) = .Cells(r, 1)

        For c = sc To Lc
            
            'Debug.Print .Cells(2, c)
            'If Not .Cells(2, c) Like "*(*" Then

            coll.Add .Cells(r, c).Value
        
            'End If
        
        Next
        
        ReDim ContentArr(1 To coll.Count)
        
        For k = 1 To coll.Count
        
            ContentArr(k) = coll.Item(k)
        
        Next

        TableData(j, 2) = ContentArr
        
        Set coll = New Collection
        
        j = j + 1

    Next
    
    Call GetLocArrInTableData '���oLocArr

End With

End Sub

Private Sub GetLocArrInTableData()

LocArr = TableData(1, 2) '�θ��u���o��

For i = 1 To UBound(LocArr)

    If LocArr(i) Like "*+*" Then LocArr(i) = mymath.TranLoc(LocArr(i))

Next

End Sub

Private Sub GetControlData() '���o�]�p�_�l�I�����k��

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
        
            collwhere.Add .Cells(sr, c)
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

MsgBox "�в��ʦ�AutoCAD�I���a�_�����ϥͦ���m�����U���I"

ret2 = MyACAD.GetPoint("�I���a�_�����ϥͦ���m�����U���I")

viewpoint(0) = ret2(0)
viewpoint(1) = ret2(1)

'Xscale = InputBox("�а�X�b��Ҭ�" & vbCrLf & "1:")
'Yscale = InputBox("�а�Y�b��Ҭ�" & vbCrLf & "1:")

End Sub

Private Sub CreateHeightLayer()

For i = 1 To UBound(AllHeightData())

    HeightName = AllHeightData(i, 0)
    
    MyACAD.acaddoc.Layers.Add (HeightName)
    
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
    
    o.Layer = "TABLE" '"�a�_��-�ϱa��"
    
    Ldpt(0) = X1 - Interval - TitleWidth
    rupt(0) = X1 - Interval
    
    Set o = MyACAD.PlotRec(Ldpt, rupt)
    
    o.Layer = "TABLE" ' "�a�_��-�ϱa��"

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
        lineObj.Layer = "TABLE" ' "�a�_��-�ϱa��"
        
        'Set xlineobj = MyACAD.AddXLine(spt, ept)
        'xlineobj.Layer = "XLINE"
    
    Next

Next

End Sub

Private Sub FillInTableTitle() '��ưϱa�����Y

Dim txtpt(2) As Double

X1 = viewpoint(0)
Y1 = viewpoint(1)

For r = 1 To TableRow

    Y1 = viewpoint(1) + (Interval + VHeight) * (r - 1)

    TableTitle = TableData(r, 1)
    
    txtpt(0) = X1 - Interval - TitleWidth / 2
    txtpt(1) = Y1 + VHeight / 2

    Set txtobj = PlotTextCenter(TableTitle, txtpt, 4 / 3 * txtheight, 2)
    
    txtobj.Layer = "TEXT" ' "�a�_��-�ϱa��r"
    
    Call MyACAD.FitTextWidth(txtobj, TitleWidth)

Next

End Sub

Sub FillInTable() '��ưϱa�����

Dim txtpt(2) As Double

Call FillInTableTitle

X1 = viewpoint(0) + startInterval
Y1 = viewpoint(1)

For r = 1 To TableRow - 1

    Y1 = viewpoint(1) + (Interval + VHeight) * (r - 1)

    ContentData = TableData(r, 2)
    'OriginLocArr = TableData(1, 2)
    
    For i = 1 To UBound(LocArr)
        
        'OriginLoc = OriginLocArr(I)
        
        If Not OriginLoc Like "*(*" Then
        
            Distance = (LocArr(i) - LocArr(1)) * 1000 / Xscale
            
            If i > 1 Then NextDistance = (LocArr(i) - LocArr(i - 1)) * 1000 / Xscale
            
            txtpt(0) = X1 + Distance
            txtpt(1) = Y1 + VHeight / 2
            
            s = ContentData(i)
            
            If IsMoved = True Then NextDistance = NextDistance - txtheight
            
            If i > 1 And NextDistance < txtheight Then
                txtpt(0) = txtpt(0) + txtheight
                IsMoved = True
            Else
                IsMoved = False
                
            End If
            
            If s Like "*(*" Then
            
                Set txtobj = PlotTextCenter(s, txtpt, 0.8 * txtheight, 2)
            
            Else
                
                If InStr(1, s, ".") <> 0 Then s = Format(s, "0.00")
                
                Set txtobj = PlotTextCenter(s, txtpt, txtheight, 2)
    
            End If
            
            txtobj.rotate MyACAD.tranPoint(txtpt), 3.14 / 2
            
            txtobj.Layer = "TEXT" ' "�a�_��-�ϱa��r"
        
        End If
        
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

'�o�̤�����

'If NeedStep < 6 Then
    
'    msg = MsgBox("�ЦA�����@�Ӥj�@�I����Ҥ�" & vbCrLf & "�ثe��Ҥج��G 1:" & Yscale & "���I�Ӥp�F!!�u��" & NeedStep & "�����I��", vbYesNo)
'    If msg = vbYes Then Yscale = InputBox("�а�Y�b��Ҭ�" & vbCrLf & "1:")

'ElseIf NeedStep > 12 Then

'    msg = MsgBox("�ЦA�����@�Ӥp�@�I����Ҥ�" & vbCrLf & "�ثe��Ҥج��G 1:" & Yscale & "���I�Ӥj�F!!�@��" & NeedStep & "�����I�h", vbYesNo)
'    If msg = vbYes Then Yscale = InputBox("�а�Y�b��Ҭ�" & vbCrLf & "1:")

'End If

NeedStep = Int(((Max - Min) * 1000 / Yscale) / BarStepHeight)

With MyACAD 'controlobject

    For i = 1 To NeedStep
    
        spt(0) = BaseHeightPoint(0)
        spt(1) = BaseHeightPoint(1) + i * BarStepHeight
        
        ept(0) = spt(0) - 15
        ept(1) = spt(1)
        
        Set lineObj = .AddLine(spt, ept)
    
        lineObj.Layer = "BAR" ' "�a�_��-���{��"
        
        'Set xlineobj = .AddXLine(spt, ept)
        'xlineobj.Layer = "XLINE"
    
        txtpt(0) = BaseHeightPoint(0) - 20
        txtpt(1) = BaseHeightPoint(1) + i * BarStepHeight
    
        s = Min + i * Yscale / (1000 / BarStepHeight)
    
        Set txtobj = PlotTextCenter(s, txtpt, txtheight, 2)
        
        txtobj.Layer = "BAR" ' "�a�_��-���{��"
        
        
    Next
    
    ept(0) = spt(0): ept(1) = spt(1)
    spt(0) = spt(0): spt(1) = BaseHeightPoint(1) + BarStepHeight
    
    Set lineObj = .AddLine(spt, ept)
    
    lineObj.Layer = "BAR" ' "�a�_��-���{��"
    
    For i = 1 To NeedStep - 1
    
        spt(0) = BaseHeightPoint(0)
        spt(1) = BaseHeightPoint(1) + (i + 0.5) * BarStepHeight
        
        ept(0) = spt(0) - 6
        ept(1) = spt(1)
        
        Set lineObj = .AddLine(spt, ept)
        lineObj.Layer = "BAR" ' "�a�_��-���{��"
    
        txtpt(0) = BaseHeightPoint(0) - 12
        txtpt(1) = BaseHeightPoint(1) + i * 10 + 5
    
        s = Min + (i + 0.5) * Yscale / (1000 / BarStepHeight)
    
        Set txtobj = PlotTextCenter(s, txtpt, txtheight, 2)
        txtobj.Layer = "BAR" ' "�a�_��-���{��"
    
    Next

End With

End Sub

Sub DrawHeight() 'ø�sPOLYLINE

Dim vertices() As Double
Dim collLoc As New Collection
Dim collHeight As New Collection

X1 = viewpoint(0) + startInterval
Y1 = viewpoint(1)

With MyACAD.acaddoc

    For i = 1 To UBound(AllHeightData)
    
        HeightName = AllHeightData(i, 0)
        HeightData = AllHeightData(i, 1)
        '-----20210706add-----
        
        .Layers.Add ("�a�_��-" & HeightName)
        
        '---------------------
        
        .ActiveLayer = .Layers("�a�_��-" & HeightName)
        
        For j = 1 To UBound(LocArr) '�o�̦�bug�A�i�H�קאּ����SE��
        
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

For j = 1 To collLoc.Count

    Distance = (collLoc.Item(j) - LocArr(1)) * 1000 / Xscale
    
    Xh = viewpoint(0) + Distance + startInterval
    Yh = BaseHeightPoint(1) + (collHeight.Item(j) - Min) * BarStepHeight * 100 / Yscale
    
    If IsShowHeightTxt = True Then
    
        txtpt(0) = Xh: txtpt(1) = Yh + 0.4
        Set txtobj = PlotTextCenter(collHeight.Item(j), txtpt, 0.4 * txtheight, 2)

        txtobj.Layer = "�a�_��-���{"

    End If

    vertices(k + 3 * (j - 1)) = Xh
    vertices(k + 1 + 3 * (j - 1)) = Yh

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

'.acaddoc.ActiveLayer = .acaddoc.Layers("�a�_��-�I�u��T")
.acaddoc.ActiveLayer = .acaddoc.Layers("L-INFO")

X1 = viewpoint(0) + startInterval
Xbar = X1
Y1 = viewpoint(1) + (TableRow - 1) * (Interval + VHeight)

For i = 1 To UBound(ControlData)

    ConLoc = mymath.TranLoc(ControlData(i, 1))
    If i <> UBound(ControlData) Then ConLocNext = TranLoc(ControlData(i + 1, 1))
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
            Set txtobj = PlotTextCenter(H, txtpt, txtheight, 2)
            Set Arc = .AddArc(cpt, VHeight / 2, -3.14 / 2, 3.14 / 2)
            
            spt(0) = cpt(0) + VHeight / 2: spt(1) = cpt(1)
            ept(0) = ept(0) + DistanceNext - VHeight / 2: ept(1) = cpt(1)
            Set lineObj = .AddLine(spt, ept)
            
        Case "C"
        
            lineObj.Delete
            txtpt(0) = cpt(0): txtpt(1) = cpt(1)
            Set txtobj = PlotTextCenter(H, txtpt, txtheight, 2)
            Set circleobj = .AddCircle(cpt, VHeight / 2)
            
            spt(0) = cpt(0) + VHeight / 2: spt(1) = cpt(1)
            ept(0) = ept(0) + DistanceNext - VHeight / 2: ept(1) = cpt(1)
            Set lineObj = .AddLine(spt, ept)
            
        Case "E"
        
            txtpt(0) = cpt(0) - VHeight / 5: txtpt(1) = cpt(1)
            Set txtobj = PlotTextCenter(H, txtpt, txtheight, 2)
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
    Set txtobj = PlotTextCenter("L=" & (ConLocNext - ConLoc) & " M,�Gh=" & Round((hNext - H), 3) & " M", txtpt, txtheight, 2)
    
    txtpt(0) = (spt(0) + ept(0)) / 2: txtpt(1) = (spt(1) + ept(1)) / 2 - VHeight / 4
    Set txtobj = PlotTextCenter("S=1:" & slope, txtpt, txtheight, 2)
    
    txtpt(0) = (spt(0) + ept(0)) / 2: txtpt(1) = (spt(1) + ept(1)) / 2 + 0.5 * VHeight + 3.5 * Interval
    Set txtobj = PlotTextCenter("�I�u�ϰ�", txtpt, 4 / 3 * txtheight, 2)
    
jump:
    
Next

End With

End Sub

Sub FillInNote() '�a�_���������Х�(�X�l)

Dim t As Integer
Dim vertices(3 * 3 - 1) As Double
Dim plineobj
Dim spt(2) As Double, ept(2) As Double, txtpt(2) As Double

X1 = viewpoint(0) + startInterval
Y1 = viewpoint(1)

Yn = TableRow * (Interval + VHeight) - Interval

On Error GoTo ERRORHANDLE

For i = 1 To UBound(NoteData)

    LocNote = TranLoc(NoteData(i, 1))
    
    Distance = (LocNote - LocArr(1)) * 1000 / Xscale
    
    Xn = X1 + Distance
    
    If i > 1 Then
        If LocNote - TranLoc(NoteData(i - 1, 1)) < 50 Then
            t = t + 1
        Else
            t = 0
        End If
    End If
    
    vertices(0) = Xn: vertices(1) = Y1 + Yn
    vertices(3) = Xn: vertices(4) = Y1 + Yn + NeedStep * 10 - t * (txtheight + Interval)
    vertices(6) = vertices(3) + 22 * Sqr(3): vertices(7) = vertices(4) + 22
    
    Set plineobj = MyACAD.AddPolyLine(vertices)
    
    plineobj.Layer = "L-INFO" '"�a�_��-�I�u��T"
    
    NoteArr = Split(NoteData(i, 2), "�B") 'NoteSeperate(NoteData(i, 2))
    
    For j = 0 To UBound(NoteArr)
    
        txtpt(0) = vertices(3) + Interval * (j - 1) + Interval
        txtpt(1) = vertices(4) + txtheight / 2 - (txtheight + Interval) * (j) + Interval
        
        s = "Sta:" & NoteData(i, 1) & ":" & NoteArr(j)
        
        Set txtobj = PlotTextCenter(s, txtpt, 4 / 3 * txtheight, 1)
        
        txtobj.rotate MyACAD.tranPoint(txtpt), 3.14 / 6
        
        txtobj.Layer = "L-INFO"  '"�a�_��-�I�u��T"
        
        If j > 0 Then
        
            spt(0) = vertices(3): spt(1) = vertices(4) - (j) * (txtheight + Interval)
            ept(0) = vertices(6): ept(1) = vertices(7) - (j) * (txtheight + Interval)
            
            Set lineObj = MyACAD.AddLine(spt, ept)
            
            lineObj.Layer = "L-INFO" ' "�a�_��-�I�u��T"
            
            t = t + 1
        
        End If
        
    Next
    
Next

ERRORHANDLE:

End Sub

Private Function FillInNoteSeperate(ByVal s As String) As Variant '�����Хܤ��j

Dim collNote As New Collection
Dim NoteArr As Variant
IsCollected = False

For i = 1 To Len(s)
    
    ch = mid(s, i, 1)
    
    NoteString = NoteString & ch
    
    If ch = "�B" Then
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

Sub TableIntroduce() '�a�_���Ϥ�ҼХ�

Dim txtpt(2)  As Double
Dim attrcoll As New Collection

txtpt(0) = viewpoint(0) - 100
txtpt(1) = viewpoint(1)

attrcoll.Add "V=1:" & Yscale
attrcoll.Add "H=1:" & Xscale

Set o = MyACAD.InsertBlock(txtpt, "LPF_SCALE", 1, attrcoll)

o.Layer = "TITLE" '"��Ҥ�"

Exit Sub

'Dim txtpt(2)  As Double

txtpt(0) = viewpoint(0) - 100
txtpt(1) = viewpoint(1) + 50

Set txtobj = PlotTextCenter("��=1:" & Xscale, txtpt, 3 * txtheight, 1)

txtpt(1) = txtpt(1) - 20

Set txtobj = PlotTextCenter("�a=1:" & Yscale, txtpt, 3 * txtheight, 1)

End Sub

Sub GetPlanDiff()

Dim collLoc As New Collection
Dim collPlanDiff As New Collection
Dim collBlank As New Collection
Dim TargetCD As String

Call GetImportantRow

With sht 'Sheets("�a�_��ø��")

    For r = 2 To .Cells(2, 1).End(xlToRight).Column
    
        collLoc.Add .Cells(2, r).Value
        'collPlanDiff.Add .Cells(rdeltaH, r).Value
        collPlanDiff.Add .Cells(rPlan, r).Value
    
    Next

End With

TargetCD = InputBox("�п�J��D�����ICD�X�W��")
'TargetType = InputBox("�п�J��D����<A,B,C,D...�ܤ@>")

With Sheets("���_��")

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
            
                .Cells(r, 4) = collPlanDiff.Item(j) & "," & TargetType(TranLoc(myloc))
                
                Exit For
            
            End If
        
        Next

        End If

    Next

Next

End With

End Sub

Function TargetType(ByVal loc As String)

With Sheets("��D������")

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
    
        tmpLoc = .Cells(2, c)
    
        tmpDiff = .Cells(rdeltaH, c)
        
        Set rng = Sheets("���߽u").Cells.Find(tmpLoc)
        
        If Not rng Is Nothing Then Sheets("���߽u").Cells(rng.Row, 4) = tmpDiff
        
    Next

End With

End Sub