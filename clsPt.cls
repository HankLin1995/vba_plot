VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================
'20210713 Author:HankLin
'store the basic information of point object
'rebuild from clsPlanMap
'define three methods to get the properties
'============================

Private PT_NUM As String
Private E As Double
Private N As Double
Private Z As Double
Private CD As String

Private CAD As New clsACAD
Private sht_PT As Object

Const NUM_DIGIT As Integer = 4

Private Sub Class_Initialize()

Set sht_PT = Sheets("總表")

End Sub

Function getTxtHeight()

txtheight = sht_PT.Range("F17")

If txtheight = 0 Then
    txtheight = CAD.GetString("請輸入文字高度(單位:mm)")
    sht_PT.Range("F17") = txtheight
End If

getTxtHeight = txtheight

End Function

Sub getPropertiesByRow(ByVal r As Integer)

With sht_PT

    PT_NUM = .Cells(r, 1)
    E = .Cells(r, 2)
    N = .Cells(r, 3)
    Z = .Cells(r, 4)
    CD = .Cells(r, 5)

End With

End Sub

Sub getPropertiesByBlock(ByVal BLK As Object)

'po = BLK.InsertionPoint
Set coll = CAD.getBlockAttr(BLK)
po = CAD.tranIPoint(BLK.InsertionPoint)

PT_NUM = coll("PT_NUM")
E = po(0)
N = po(1)
Z = coll("PT_Z")
CD = coll("PT_CD")

End Sub

Sub getPropertiesByBlock_TYLin(ByVal BLK As Object)

'po = BLK.InsertionPoint
Set coll = CAD.getBlockAttr(BLK)
po = CAD.tranIPoint(BLK.InsertionPoint)

If coll.Count = 0 Then Exit Sub

PT_NUM = coll("PT_NUM")
E = po(0)
N = po(1)
Z = coll("PT_Z")
CD = coll("PT_CD")

End Sub


Sub getPropertiesByBlock_Doris(ByVal BLK As Object)

'po = BLK.InsertionPoint
Set coll = CAD.getBlockAttr(BLK)
po = CAD.tranIPoint(BLK.InsertionPoint)

If coll.Count = 0 Then Exit Sub

PT_NUM = coll(1)
E = po(0)
N = po(1)
Z = coll(2)
CD = coll(3)

End Sub

Sub getPropertiesByBlock_Tien(ByVal BLK As Object) ', PT_NUM, E, N)

'po = BLK.InsertionPoint
Set coll = CAD.getBlockAttr(BLK)
po = CAD.tranIPoint(BLK.InsertionPoint)

If coll.Count = 0 Then Exit Sub

PT_NUM = coll(1)
E = po(0)
N = po(1)

End Sub


Sub getPropertiesByUser(ByVal bPT_NUM As String, ByVal bE As Double, ByVal bN As Double, ByVal bZ As Double, ByVal bCD As String)

PT_NUM = bPT_NUM
E = bE
N = bN
Z = bZ
CD = bCD

End Sub

Sub CreatePoint(ByVal txtheight As Double)

If txtheight = 0 Then txtheight = getTxtHeight

Dim txtpt(2) As Double

txtpt(0) = E: txtpt(1) = N

'coll_attr.Add PT_NUM
'coll_attr.Add Z
'coll_attr.Add CD

arr = Array(PT_NUM, Z, CD)

Set ptObj = CAD.InsertBlock(txtpt, "Point", txtheight, arr) ' coll_attr)

Call CAD.setLayer(ptObj, "平面圖-點")

'ptObj.Layer = "平面圖-點"

End Sub




Sub AppendData(Optional ByVal r As Double)

With sht_PT

    If r = 0 Then r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    
    .Cells(r, 1) = PT_NUM
    .Cells(r, 2) = Round(E, NUM_DIGIT)
    .Cells(r, 3) = Round(N, NUM_DIGIT)
    .Cells(r, 4) = Z
    .Cells(r, 5) = CD

End With

End Sub

Sub AppendData_TYLin(Optional ByVal r As Double)

With Sheets("FIX_POINT")

    If r = 0 Then r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    
    .Cells(r, 1) = PT_NUM
    .Cells(r, 2) = Round(E, NUM_DIGIT)
    .Cells(r, 3) = Round(N, NUM_DIGIT)
    .Cells(r, 4) = Z
    .Cells(r, 5) = CD

End With

End Sub
