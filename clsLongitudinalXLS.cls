VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLongitudinalXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private CAD As New clsACAD
Private math As New clsMath
Private collXLine As New Collection
Private shtLsec As Object
Private shtHsec As Object

Private Sub Class_Initialize()

Set shtLsec = Sheets("縱斷面")
Set shtHsec = Sheets("橫斷面")

End Sub

Sub RenewTable()

Dim colltmp As New Collection
Dim Xco As Double
Dim Yco As Double
Dim Xnum As Double
Dim Ynum As Double
Dim it As Object 'AcadLine
Dim ssetX As Object 'AcadSelectionSet
Dim ssetPL As Object 'AcadSelectionSet

MsgBox "請移動至CAD圖面依照指示辦理"

IsGetColumn = False
r = 3
Xsc = returnScale(0, Xco, Xnum)
Ysc = returnScale(1, Yco, Ynum)

CAD.acaddoc.Utility.prompt "請框選欲查詢的樁號"

Set ssetX = CAD.CreateSSET("SSETX")

CAD.acaddoc.Utility.prompt "請框選欲查詢的縱斷面高程線"

Set ssetPL = CAD.CreateSSET("SSETPL")

For Each PL In ssetPL
    
    colltmp.Add PL.Layer
    
    For Each it In ssetX
    
        ret = it.IntersectWith(PL, 1) 'acExtendThisEntity)
        
        pt = it.startpoint
        
        If UBound(ret) = 2 Then
        
            l = (ret(0) - Xco) * 2500 / 1000 + Xnum
            H = (ret(1) - Yco) * Ysc / 1000 + Ynum
            
            colltmp.Add l & "," & H
        
        End If
        
        xl = (pt(0) - Xco) * 2500 / 1000 + Xnum
        
        If IsGetColumn = False Then
        
            xl = (pt(0) - Xco) * 2500 / 1000 + Xnum: collXLine.Add xl
            
        End If
        
    Next
    
    If IsGetColumn = False Then Call getXLineColumn: IsGetColumn = True
    
    Call WriteToXLS(r, colltmp)

    r = r + 1

Next

shtLsec.Activate

End Sub

Private Sub WriteToXLS(ByVal r As Integer, ByRef colltmp As Variant)

'On Error Resume Next

shtLsec.Cells(r, 1) = colltmp(1)

For i = 2 To colltmp.Count

    tmp = Split(colltmp(i), ",")
    
    tmpLoc = tmp(0)
    H = tmp(1)
    c = collXLine(tmpLoc)
    
    shtLsec.Cells(r, c) = H

Next

Set colltmp = New Collection

End Sub

Private Sub getXLineColumn()

Dim collXLineRow As New Collection

j = 2

With shtLsec
    
    .Cells.Clear
    .Cells(1, 1) = "樁號"
    
    For i = collXLine.Count To 1 Step -1
    
        .Cells(1, j) = collXLine(i)
        
        collXLineRow.Add j, Key:=CStr(collXLine(i))
        j = j + 1
        
    Next

End With

Set collXLine = collXLineRow

End Sub

Private Function returnScale(ByVal mode As Integer, ByRef co As Double, ByRef Num As Double)

prompt = "請點選Y軸的基準點"
If mode = 0 Then prompt = "請點選X軸的基準點"

With CAD

    spt = .GetPoint(prompt)
    Num = .GetString("請輸入基準數值:")
    co = spt(mode)
    ept = .GetPoint("請點選基準長度終點")
    
    len_paper = ept(mode) - spt(mode)
    len_real = .GetString("請輸入實際長度(公尺):")
    
    returnScale = len_real / len_paper * 1000

End With

End Function

Sub ExtractToLSection()

Dim collBlank As New Collection

r = 2

With shtHsec
    
    shtLsec.Cells.ClearContents

    shtLsec.Cells(1, 1) = "樁號"
    
    For rTarget = 2 To .Range("F" & Rows.Count).End(xlUp).Row '初始化
        
        HeightName = .Range("H" & rTarget)
        TargetCol = .Range("G" & rTarget)
        shtLsec.Cells(1, TargetCol) = HeightName
        
    Next

    collBlank.Add 1
    
    For rBlank = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
    
        If .Cells(rBlank, 1) = "" Then collBlank.Add rBlank
    
    Next
    
    collBlank.Add .Cells(Rows.Count, 1).End(xlUp).Row + 1

    For i = 1 To collBlank.Count - 1 'To 1 Step -1
        
        sr = collBlank.Item(i) + 1
        er = collBlank.Item(i + 1) - 1
        
        where = math.TranLoc(.Cells(sr, 4))
        shtLsec.Cells(r, 1) = where
        
        For rr = sr To er
        
            CD = .Cells(rr, 3)
            H = Round(.Cells(rr, 2), 3)
            
            For rTarget = 2 To .Range("G" & Rows.Count).End(xlUp).Row
            
                If .Range("F" & rTarget) = CD Then
                
                    TargetCol = .Range("G" & rTarget)
                    shtLsec.Cells(r, TargetCol) = H
                    Exit For
                    
                End If
                    
            Next
        
        Next
    
        r = r + 1
    
    Next

End With

End Sub

Sub ExtractToLSectionSort()

With shtLsec

    Lc = .Cells(1, 200).End(xlToLeft).Column
    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    .Range("A2:" & Chr(Lc + 64) & lr).Sort key1:=.Range("A2:A" & lr), order1:=xlAscending
    .Range("A1:" & Chr(Lc + 64) & lr).Copy
    .Range(Chr(Lc + 65) & "1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    .Columns("A:" & Chr(Lc + 64)).Delete
    .Activate

End With

End Sub
