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
Private math As New clsMath
Private pt As New clsPt 'new object 20210713
Private myFunc As New clsMyfunction

Const NUM_DIGIT As Integer = 3

Private Sub Class_Initialize()

Set sht = Sheets("�`��")

End Sub

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


'Function CollPointTable(ByVal findkey As String) '�I��ƹ����Ҳ�
'
''�p�G�٨S�إ�collPtTable�h���i��Ĥ@��������
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

'===========�°w��EXCEL���B�z=======================

Sub RotatePoint(Optional ByVal IsAsk = True, Optional ByVal X0 As Double, Optional ByVal Y0 As Double, Optional ByVal ptang As Double)

On Error GoTo ERRORHANDLER

With sht

lr = .Cells(Rows.Count, 1).End(xlUp).Row

If IsAsk = True Then

    ptnum = InputBox("�п�J�������I�Ǹ�(PT_NUM)")
    ptang = CDbl(InputBox("�п�J���ਤ��(degree)"))
    
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

    fi = math.getAz(X0, Y0, X, Y)
    l = Sqr((X - X0) ^ 2 + (Y - Y0) ^ 2)
    
    thida = fi + ptang

    dx = l * math.degcos(thida)
    dy = l * math.degsin(thida)

    newx = X0 + dx
    newy = Y0 + dy

    .Cells(r, "B") = newx
    .Cells(r, "C") = newy

Next

End With

Exit Sub

ERRORHANDLER:

MsgBox "��J���ئ��~!", vbCritical

End Sub

Sub MovePoint(Optional ByVal IsAsk As Boolean = True, Optional ByVal deltaX As Double, Optional ByVal deltaY As Double)

With sht

    If IsAsk = True Then

        deltaX = Val(InputBox("�п�JX�b�첾�Z��"))
        deltaY = Val(InputBox("�п�JY�b�첾�Z��"))
    
    End If
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        .Cells(r, 2) = .Cells(r, 2) + deltaX
        .Cells(r, 3) = .Cells(r, 3) + deltaY
    
    Next

End With

End Sub

Sub AlignPoint()

Call getXYfromTable(Xs, Ys, "�Ĥ@�I�I�Ǹ�(PT_NUM)")

Xe = Val(InputBox("�Ĥ@�IX�b��:" & Xs & vbCrLf & "������ؼ��IX="))
Ye = Val(InputBox("�Ĥ@�IY�b��:" & Ys & vbCrLf & "������ؼ��IY="))

Call getXYfromTable(Xss, Yss, "�ĤG�I�I�Ǹ�(PT_NUM)")

Xee = Val(InputBox("�ĤG�IX�b��:" & Xss & vbCrLf & "������ؼ��IX="))
Yee = Val(InputBox("�ĤG�IY�b��:" & Yss & vbCrLf & "������ؼ��IY="))

dx = Xe - Xs
dy = Ye - Ys

Call MovePoint(False, dx, dy)

fi1 = math.getAz(Xe, Ye, Xss + dx, Yss + dy)
fi2 = math.getAz(Xe, Ye, Xee, Yee)

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

Sub ExportCSV() '��XCSV

Set wk = Workbooks.Add

With sht

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    wk.Sheets(1).Cells(r - 1, 1) = .Cells(r, 2)
    wk.Sheets(1).Cells(r - 1, 2) = .Cells(r, 3)

Next

f = Application.GetSaveAsFilename & ".csv"

If f = "False.csv" Then MsgBox "���|���s��": wk.Close False: Exit Sub

wk.SaveAs Filename:=f, FileFormat:=xlCSV, CreateBackup:=False

wk.Close False

End With

End Sub

Sub ImportTXT() '��zTXT

mode = InputBox("�q�����U���᪺�ƧǶ���" & vbCrLf & "1.P,E,N,Z,CD" & vbCrLf & "2.P,N,E,Z,CD")

With sht

    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        mydata = .Cells(r, 1)
        
        tmp = Split(mydata, ",")
        
        For i = LBound(tmp) To UBound(tmp)
        
            Select Case i
            
            Case 0: c = 1
            Case 1
            
                c = 2
                If mode = 2 Then c = 3
                
            Case 2
            
                c = 3
                If mode = 2 Then c = 2
                
            Case 3: c = 4
            Case 4: c = 5
            
            End Select
        
            .Cells(r, c) = tmp(i)
        
        Next
    
    Next

End With

End Sub

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