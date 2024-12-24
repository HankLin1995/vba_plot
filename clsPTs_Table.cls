VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPTs_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================
'20210713 Author:HankLin
'for those data only used in EXCEL
'move from clsPlanMap
'============================

Private collPtTable As New Collection '20210615

Private sht As Object
Private Math As New clsMath
Private pt As New clsPt 'new object 20210713
Private myFunc As New clsMyfunction

Const NUM_DIGIT As Integer = 3

Private Sub Class_Initialize()

Set sht = Sheets("總表")

End Sub

Function getNearlyPointHeight(ByVal X1, ByVal Y1, ByVal diff, ByRef nearX, ByRef nearY)

With sht

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        X0 = .Cells(r, 2)
        Y0 = .Cells(r, 3)
        Z = .Cells(r, 4)
        distance = Math.getLengthCO(X0, Y0, X1, Y1)
        
        If distance < diff Then
            nearX = X0
            nearY = Y0
            getNearlyPointHeight = Z
            Exit For
        End If
    
    Next

End With

End Function

Function getCollPointTable()

Dim coll As New Collection

With sht

    For r = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
    
        mykey = .Cells(r, 2) & ":" & .Cells(r, 3)
    
        On Error Resume Next
    
        coll.Add .Cells(r, 4).Value, mykey
        
        On Error GoTo 0
    
    Next

End With

Set getCollPointTable = coll

End Function

Function getCollPointTable_TYLin()

Dim coll As New Collection

With Sheets("FIX_POINT")

    For r = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row

        mykey = .Cells(r, 2) & ":" & .Cells(r, 3)

        On Error Resume Next

        coll.Add .Cells(r, 4).Value, mykey

        On Error GoTo 0

    Next

End With

Set getCollPointTable_TYLin = coll

End Function


'Function CollPointTable(ByVal findkey As String) '點資料對應模組
'
''如果還沒建立collPtTable則先進行第一次的對應
'
'If collPtTable.Count = 0 Then Set collPtTable = getCollPointTable
'
'On Error GoTo ERRORHANDLE
'
'CollPointTable = collPtTable(findkey)
'
'Exit Function
'
'ERRORHANDLE:
'
'CollPointTable = 0
'
'End Function

'===========純針對EXCEL做處理=======================

Sub RotatePoint(Optional ByVal IsAsk = True, Optional ByVal X0 As Double, Optional ByVal Y0 As Double, Optional ByVal ptang As Double)

On Error GoTo ERRORHANDLER

With sht

lr = .Cells(Rows.Count, 1).End(xlUp).Row

If IsAsk = True Then

    ptnum = InputBox("請輸入旋轉基準點序號(PT_NUM)")
    ptang = CDbl(InputBox("請輸入旋轉角度(degree)"))
    
    For r = 2 To lr
    
        If .Cells(r, 1) = Val(ptnum) Then
        
            X0 = .Cells(r, 2)
            Y0 = .Cells(r, 3)
            
            Exit For
            
        End If
    
    Next

End If

For r = 2 To lr

    X = .Cells(r, 2)
    Y = .Cells(r, 3)

    fi = Math.getAz(X0, Y0, X, Y)
    L = Sqr((X - X0) ^ 2 + (Y - Y0) ^ 2)
    
    thida = fi + ptang

    dx = L * Math.degcos(thida)
    dy = L * Math.degsin(thida)

    newx = X0 + dx
    newy = Y0 + dy

    .Cells(r, "B") = newx
    .Cells(r, "C") = newy

Next

End With

Exit Sub

ERRORHANDLER:

MsgBox "輸入項目有誤!", vbCritical

End Sub

Sub MovePoint(Optional ByVal IsAsk As Boolean = True, Optional ByVal deltaX As Double, Optional ByVal deltaY As Double)

With sht

    If IsAsk = True Then

        deltaX = Val(InputBox("請輸入X軸位移距離"))
        deltaY = Val(InputBox("請輸入Y軸位移距離"))
    
    End If
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        .Cells(r, 2) = .Cells(r, 2) + deltaX
        .Cells(r, 3) = .Cells(r, 3) + deltaY
    
    Next

End With

End Sub

Sub AlignPoint()

Call getXYfromTable(Xs, Ys, "第一點點序號(PT_NUM)")

Xe = Val(InputBox("第一點X軸為:" & Xs & vbCrLf & "對應於目標點X="))
Ye = Val(InputBox("第一點Y軸為:" & Ys & vbCrLf & "對應於目標點Y="))

Call getXYfromTable(Xss, Yss, "第二點點序號(PT_NUM)")

Xee = Val(InputBox("第二點X軸為:" & Xss & vbCrLf & "對應於目標點X="))
Yee = Val(InputBox("第二點Y軸為:" & Yss & vbCrLf & "對應於目標點Y="))

dx = Xe - Xs
dy = Ye - Ys

Call MovePoint(False, dx, dy)

fi1 = Math.getAz(Xe, Ye, Xss + dx, Yss + dy)
fi2 = Math.getAz(Xe, Ye, Xee, Yee)

deltafi = fi2 - fi1

Call RotatePoint(False, Xe, Ye, deltafi)

End Sub

Private Sub getXYfromTable(ByRef X, ByRef Y, ByVal prompt)

ptnum = InputBox(prompt)

With sht

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lr
    
        If .Cells(r, 1) = Val(ptnum) Then
        
            X = .Cells(r, 2)
            Y = .Cells(r, 3)
            
            Exit For
            
        End If
    
    Next
    
End With
End Sub

Sub ExportCSV_ForWEBGIS() '輸出CSV

Set wk = Workbooks.Add

With sht

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    wk.Sheets(1).Cells(r - 1, 1) = .Cells(r, 2)
    wk.Sheets(1).Cells(r - 1, 2) = .Cells(r, 3)

Next

f = Application.GetSaveAsFilename & ".csv"

If f = "False.csv" Then MsgBox "文件尚未存檔": wk.Close False: Exit Sub

wk.SaveAs Filename:=f, FileFormat:=xlCSV, CreateBackup:=False

wk.Close False

End With

End Sub

Sub ExportCSV2() '輸出CSV,0511add,ForSTONEX

Set wk = Workbooks.Add

With sht

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

wk.Sheets(1).Cells(1, 1) = "點名"
wk.Sheets(1).Cells(1, 2) = "北坐標"
wk.Sheets(1).Cells(1, 3) = "東坐標"
wk.Sheets(1).Cells(1, 4) = "高程"
wk.Sheets(1).Cells(1, 5) = "編碼"

For r = 2 To lr

    wk.Sheets(1).Cells(r, 1) = .Cells(r, 1)
    wk.Sheets(1).Cells(r, 2) = .Cells(r, 3)
    wk.Sheets(1).Cells(r, 3) = .Cells(r, 2)
    wk.Sheets(1).Cells(r, 4) = .Cells(r, 4)
    wk.Sheets(1).Cells(r, 5) = .Cells(r, 5)

Next

f = Application.GetSaveAsFilename

If f = "False.csv" Then MsgBox "文件尚未存檔": wk.Close False: Exit Sub

wk.SaveAs Filename:=f & "csv", FileFormat:=xlCSV, CreateBackup:=False

wk.Close False

End With

End Sub

Sub ExportCSV() 'for frank

Set wk = Workbooks.Add

With sht

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    wk.Sheets(1).Cells(r - 1, 1) = .Cells(r, 5)
    wk.Sheets(1).Cells(r - 1, 2) = .Cells(r, 3)
    wk.Sheets(1).Cells(r - 1, 3) = .Cells(r, 2)

Next

f = Application.GetSaveAsFilename & ".csv"

If f = "False.csv" Then MsgBox "文件尚未存檔": wk.Close False: Exit Sub

wk.SaveAs Filename:=f, FileFormat:=xlCSV, CreateBackup:=False

wk.Close False

End With

End Sub

'Sub ImportTXT() '整理TXT
'
'mode = InputBox("從機器下載後的排序順序" & vbCrLf & "1.P,E,N,Z,CD" & vbCrLf & "2.P,N,E,Z,CD")
'
'With sht
'
'    lr = .Cells(Rows.Count, 1).End(xlUp).Row
'
'    For r = 2 To lr
'
'        mydata = .Cells(r, 1)
'
'        tmp = Split(mydata, ",")
'
'        For i = LBound(tmp) To UBound(tmp)
'
'            Select Case i
'
'            Case 0: c = 1
'            Case 1
'
'                c = 2
'                If mode = 2 Then c = 3
'
'            Case 2
'
'                c = 3
'                If mode = 2 Then c = 2
'
'            Case 3: c = 4
'            Case 4: c = 5
'
'            End Select
'
'            .Cells(r, c) = tmp(i)
'
'        Next
'
'    Next
'
'End With
'
'End Sub

Sub test_GetOldPTs()

Dim CAD As New clsACAD
Dim collPT_NUM As New Collection
Dim collN As New Collection
Dim collE As New Collection
Dim collZ As New Collection
Dim collCD As New Collection

Set sset = CAD.CreateSSET

For Each it In sset

    Debug.Print it.Layer
    
    Select Case it.Layer
    
    Case "PT_Point"
        co = it.coordinates
        collE.Add co(0)
        collN.Add co(1)
        
    Case "PT_NUM": collPT_NUM.Add it.TextString
    
    Case "PT_Z":  collZ.Add it.TextString
    
    Case "PT_CD": collCD.Add it.TextString
    
    End Select

Next

With sht

    r = 2
    
    For i = 1 To collE.Count
        
       .Cells(r, 1) = collPT_NUM(i)
       .Cells(r, 2) = collE(i)
       .Cells(r, 3) = collN(i)
       .Cells(r, 4) = collZ(i)
       .Cells(r, 5) = collCD(i)
        r = r + 1
    Next

End With

End Sub
