VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimpleHSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myloc As String
Private SHsht As Object
Private Hsht As Object
Private ShapeArr As Variant
Private collLoc As New Collection
Private ACAD As Object
Private mo As Object
Private MyACAD As New clsACAD
Private math As New clsMath
Private txtheight As Double

Public myUnit As String
Public myWidth As Double
Public TableRow As Integer

Private Sub Class_Initialize()

Set SHsht = Sheets("標竿橫斷")
Set Hsht = Sheets("橫斷面")

Set ACAD = MyACAD.acaddoc   'acadapp.ActiveDocument
Set mo = ACAD.ModelSpace

ACAD.Layers.Add "PT_Point"
ACAD.Layers.Add "PT_NUM"
ACAD.Layers.Add "PT_CD"
ACAD.Layers.Add "PT_Z"

End Sub

Sub ReadSHTable() 'ByVal r As Integer)

Dim collX As New Collection
Dim collZ As New Collection
Dim collCD As New Collection

If myUnit = "" Then myUnit = "公尺" 'InputBox("請輸入紅白標竿記錄表的單位為 公尺 or 公分", , "公尺")

Select Case myUnit

Case "公尺": tr = 1
Case "公分": tr = 100

End Select

r = TableRow

With SHsht
    
    Lc = .Cells(r, Columns.Count).End(xlToLeft).Column
    
    myloc = .Cells(r, 1)

    tmp = Split(.Cells(r + 2, 1), ",")

    lcol = Asc(tmp(0)) - 64
    rcol = Asc(tmp(1)) - 64
    
    'tmp2 = Split(.Cells(r + 1, 1), ",")
    
    z_cl = .Cells(r + 1, 1) 'Val(tmp2(1))
    
    For c = 2 To Lc
    
        sumX = myWidth / tr
        sumZ = z_cl
        myCD = .Cells(r + 2, c)
        'myCD_cl = tmp2(0)
        
        If myCD = "" Then myCD = "NONE"
        
        If c = rcol Then
            
            collX.Add sumX
            collZ.Add sumZ
            collCD.Add "中心樁" 'myCD_cl
            
        End If
    
        If c <= lcol Then
        
            For tmpc = c To lcol
                
                dx = .Cells(r, tmpc)
                dz = .Cells(r + 1, tmpc)
                
                sumX = sumX - dx / tr
                sumZ = sumZ - dz / tr
            
            Next
        
        Else
        
            For tmpc = rcol To c
            
                dx = .Cells(r, tmpc)
                dz = .Cells(r + 1, tmpc)
                
                sumX = sumX + dx / tr
                sumZ = sumZ + dz / tr
            
            Next
        
        End If
        
        collX.Add sumX
        collZ.Add sumZ
        collCD.Add myCD
    
    Next
    
    ReDim ShapeArr(1 To collX.Count, 1 To 3)
    
    For i = 1 To collX.Count
        
        ShapeArr(i, 1) = collX(i)
        ShapeArr(i, 2) = collZ(i)
        ShapeArr(i, 3) = collCD(i)
        
    Next

End With

End Sub

Sub SHTable2Hsection()

With Hsht

lr = .Cells(Rows.Count, 1).End(xlUp).Row + 2

If lr = 3 Then lr = 2

.Cells(lr, 4) = myloc

For i = 1 To UBound(ShapeArr, 1)
    
    For c = 1 To 3
    
        .Cells(lr + j, c) = ShapeArr(i, c)
    
    Next
    
    j = j + 1

Next

.Activate

End With

End Sub

Sub ClearHSection() '清空橫斷面資料表

With Hsht

    r = 2
    
    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    .Range("A" & r).Resize(lr, 4).ClearContents

End With

End Sub

Sub CollectHsec()

Dim arr() As Variant

With Hsht

Set collb = GetBlankRow(Hsht, 2)

For i = 1 To collb.Count - 1

    sr = collb.Item(i) + 1
    er = collb.Item(i + 1) - 1
    key_loc = .Cells(sr, 4)
    
    ReDim arr(1 To er - sr + 1, 1 To 3)
    j = 0
    
    For r = sr To er
    
        j = j + 1
        arr(j, 1) = .Cells(r, 1): arr(j, 2) = .Cells(r, 2): arr(j, 3) = .Cells(r, 3)
        
    Next
    
    collLoc.Add arr, key_loc

Next

End With

End Sub

Sub GetCADLoc() 'CAD圖面依照橫斷面資料表展點

Dim xdataOut As Variant
Dim xtypeOut As Variant

Set sset = CreateSSET(3)

For Each it In sset

    Dim arr As Variant

    If TypeName(it) = "IAcadLine" Then

        it.GetXData "", xtypeOut, xdataOut
        On Error Resume Next
        arr = collLoc(xdataOut(0))
        Debug.Print Err.Number

        If Err.Number = 5 Then GoTo errhandle

        arr = collLoc(xdataOut(0))
    
        spt = it.startpoint
        ept = it.endpoint
        
        fi = math.getAz(spt(0), spt(1), ept(0), ept(1))
    
        For i = LBound(arr, 1) To UBound(arr, 1)
        
            ptWidth = arr(i, 1)
            Z = arr(i, 2)
            CD = arr(i, 3)
            X = spt(0) + ptWidth * Sin(fi / 180 * 4 * Atn(1))
            Y = spt(1) + ptWidth * Cos(fi / 180 * 4 * Atn(1))
            PT_NUM = PT_NUM + 1
            
            args = Array(X, Y, Z, PT_NUM, CD)
            Call CreatePointData(args)
            
        Next
    
    End If
errhandle:
Next

End Sub

Sub CreatePointData(ByVal args As Variant)

Dim txtpt(2) As Double

If txtheight = 0 Then txtheight = Val(InputBox("請輸入文字高度(單位:mm)"))

X = args(0)
Y = args(1)
Z = args(2)
PT_NUM = args(3)
CD = args(4)

txtpt(0) = X: txtpt(1) = Y

ACAD.ActiveLayer = ACAD.Layers("PT_Point")
Set ptObj = mo.AddPoint(txtpt)

ACAD.ActiveLayer = ACAD.Layers("PT_NUM")
Set PT_NUMtxtobj = mo.AddText(PT_NUM, txtpt, txtheight)

txtpt(1) = txtpt(1) - txtheight

ACAD.ActiveLayer = ACAD.Layers("PT_CD")
Set CDtxtobj = mo.AddText(CD, txtpt, txtheight)
ACAD.ActiveLayer.Color = 6
txtpt(1) = txtpt(1) - txtheight

ACAD.ActiveLayer = ACAD.Layers("PT_Z")
Set Ztxtobj = mo.AddText(Z, txtpt, txtheight)

End Sub

Function GetBlankRow(ByVal sht As Object, ByVal sr As Integer, Optional sc As Integer = 1) As Collection

Dim collBlank As New Collection

With sht

    collBlank.Add sr - 1
    
    For r = sr To .Cells(.Rows.Count, 1).End(xlUp).Row
    
        If .Cells(r, 1) = "" Then collBlank.Add r
    
    Next
    
    collBlank.Add .Cells(Rows.Count, 1).End(xlUp).Row + 1

End With

Set GetBlankRow = collBlank

End Function



