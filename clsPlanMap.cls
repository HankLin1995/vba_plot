VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlanMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================
'20210713 Author:HankLin
'for those data only used to CAD
'rebuild clsPlanMap
'展點、匯入點、整理數據、畫線
'============================

Private sht As Object
Private CAD As New clsACAD
Private math As New clsMath
Private pt As New clsPt 'new object 20210713
Private myFunc As New clsMyfunction 'new object 20210713

Private IsLT As Boolean

Private Sub Class_Initialize()

Set sht = Sheets("總表")

IsLT = False
If sht.optNOVBA = True Then IsLT = True

End Sub

Sub ImportDataToCAD() 'DrawPoint

If IsLT = True Then MsgBox "LT版本不支援喔!": Exit Sub

txtheight = Val(InputBox("請輸入文字高度(單位:mm)", , 0.5))

Call checkBlockExist("Point")

With sht

lr = .Cells(1, 1).End(xlDown).Row

    For r = 2 To lr
    
        Call pt.getPropertiesByRow(r)
        Call pt.CreatePoint(txtheight)
    
    Next

.Range("F17") = txtheight '儲存資訊於此供未來應用

End With

'ACAD.SendCommand "zoom" & vbCr & "e" & vbCr

End Sub

Sub CreatePointByUser()

With sht

    myans = "Y"
    
    Do While myans = "Y"
    
        mypt = CAD.GetPoint("Select the new point")
        
        E = Round(mypt(0), 3)
        N = Round(mypt(1), 3)
        
        PT_NUM = CAD.GetString("PT_Num=?")
        Z = CAD.GetString("Z=?")
        CD = CAD.GetString("CD=?")
        
        Call pt.getPropertiesByUser(PT_NUM, E, N, Z, CD)
        Call pt.CreatePoint(txtheight)
        Call pt.AppendData
        
        myans = CAD.GetString("繼續補點嗎?(Y/N)")
        
    Loop
    
End With

End Sub

Sub ExportDataToExcel() '匯入點

If IsLT = True Then MsgBox "LT版本不支援喔!": Exit Sub

If MsgBox("是否要清除既有點資料?", vbYesNo) = vbYes Then Call myFunc.ClearData(sht, 2, 1, 5)

MsgBox "請移動至CAD選取要匯入的點位"

Set sset = CAD.CreateSSET("SS1")

For Each it In sset

    Debug.Print TypeName(it)

        If TypeName(it) Like "*Block*" Then
    
        Call pt.getPropertiesByBlock(it)
        Call pt.AppendData
    
    End If

Next

End Sub

Sub ExportDataToExcel_TYLin() '匯入點

If IsLT = True Then MsgBox "LT版本不支援喔!": Exit Sub

If MsgBox("是否要清除既有點資料?", vbYesNo) = vbYes Then Call myFunc.ClearData(Sheets("FIX_POINT"), 2, 1, 5)

MsgBox "請移動至CAD選取要匯入的點位"

Set sset = CAD.CreateSSET("SS1")

For Each it In sset

        If TypeName(it) Like "*Block*" Then
    
            Debug.Print TypeName(it)
    
        Call pt.getPropertiesByBlock_TYLin(it)
        Call pt.AppendData_TYLin
    
    End If

Next

End Sub

Sub ExportDataToExcel_Doris() '匯入點

If IsLT = True Then MsgBox "LT版本不支援喔!": Exit Sub

If MsgBox("是否要清除既有點資料?", vbYesNo) = vbYes Then Call myFunc.ClearData(Sheets("FIX_POINT"), 2, 1, 5)

MsgBox "請移動至CAD選取要匯入的點位"

Set sset = CAD.CreateSSET("SS1")

For Each it In sset

        If TypeName(it) Like "*Block*" Then
    
            Debug.Print TypeName(it)
    
        Call pt.getPropertiesByBlock_Doris(it)
        Call pt.AppendData_TYLin
    
    End If

Next

End Sub

Sub ExportDataToExcel_OldPT() '匯入點(舊值)

Dim PTsObj As New clsPTs_Table

If IsLT = True Then MsgBox "LT版本不支援喔!": Exit Sub

If MsgBox("是否要清除既有點資料?", vbYesNo) = vbYes Then Call myFunc.ClearData(sht, 2, 1, 5)

MsgBox "請移動至CAD選取要匯入的點位"

Call PTsObj.test_GetOldPTs

End Sub


Sub ExportPLToExcel() '匯入線

If IsLT = True Then MsgBox "LT版本不支援喔!": Exit Sub

If MsgBox("是否要清除既有線資料?", vbYesNo) = vbYes Then
    Call myFunc.ClearData(Sheets("特徵線"), 2, 1, 3)
End If

MsgBox "請移動至CAD選取要匯入的線"

'Dim CAD As New clsACAD
Set PLs = CAD.CreateSSET("SS1")

For Each PL In PLs

    Dim obj As New clsPL
    Call obj.getPropertiesByPL(PL)
    Call obj.AppendData

Next

End Sub

'========整理數據===============

Sub ReArrangeCD(ByVal mode As Byte) '整理數據排序

Call myFunc.ClearData(sht, 2, 7, 6) 'r=7,xrange=6
Call myFunc.ClearData(sht, 2, 14, 1) 'r=14,xrange=1

Set collCD = CollectCD

With sht

lr = .Cells(1, 1).End(xlDown).Row

rData = 2
rDefault = 2

Select Case mode

Case 1: s = "H"
Case 2: s = "I"
Case 3: s = "G"

End Select

For i = 1 To collCD.Count

    sr = rData
    
    For r = 2 To lr
    
        TargetCD = CStr(.Cells(r, 5))
        
        If TargetCD = collCD(i) Then
            
            For c = 1 To 5
            
                .Cells(rData, 6 + c) = .Cells(r, c)
            
            Next
        
            rData = rData + 1
        
        End If
    
    Next
    
    rData = rData + 1

    er = rData - 1
    
    .Range("G" & sr & ":K" & er).Sort key1:=Range(s & sr & ":" & s & er), order1:=xlAscending
    'Debug.Print "sr:" & sr & " er:" & er
    
    If er - sr > 1 Then
        .Cells(rDefault, "N") = collCD(i)
        rDefault = rDefault + 1
    End If
    
Next

End With

End Sub

Sub SetDefaultFeature() '20200729 設定Feature

Set collDefaultCD = CollectDefaultCD

With sht

    Set collBlank = myFunc.getBlankColl(sht, 7)
    
    For i = 1 To collBlank.Count - 1
    
        sr = collBlank.Item(i) + 1
        er = collBlank.Item(i + 1) - 1
        CD = .Cells(sr, 11)
        ptnum = er - sr + 1
        
        If ptnum > 1 Then
        
            On Error Resume Next
            .Cells(sr, "L") = collDefaultCD(CD)
            On Error GoTo 0
            
        End If
        
    Next

End With

End Sub

Private Function CollectCD()

Dim coll As New Collection

With sht

    lr = .Cells(1, 1).End(xlDown).Row
    
    For r = 2 To lr
    
        TargetCD = Trim(.Cells(r, 5))
        
        On Error Resume Next: coll.Add TargetCD, TargetCD: On Error GoTo 0
    
    Next

End With

Set CollectCD = coll

End Function

Private Function CollectDefaultCD() '20200728

Dim coll As New Collection

With sht
    
    lr = .Cells(.Rows.Count, "N").End(xlUp).Row
    
    For r = 2 To lr
        
        DefaultCD = Trim(.Cells(r, "N"))
        DefaultFeature = .Cells(r, "O")
        
        coll.Add DefaultFeature, DefaultCD
        
    Next
    
End With

Set CollectDefaultCD = coll

End Function

'Private Sub ClearData(ByVal sr As Integer, ByVal col As Integer, ByVal xrange As Integer)
'
'With sht
'
'lr = .Cells(.Rows.Count, col).End(xlUp).Row
'.Cells(sr, col).Resize(lr, xrange).ClearContents
'
'End With
'
'End Sub

'=============畫線=============

Sub DrawLine(ByVal Radius As Double)

If IsLT = True Then MsgBox "LT版本不支援喔!": Exit Sub

'Dim plineobj As Object 'AcadPolyline
'Dim circleobj As Object 'AcadCircle
Dim vertices() As Double
Dim cpt(2) As Double

With sht
    
    Set collBlank = myFunc.getBlankColl(sht, 7)

    For i = 1 To collBlank.Count - 1
    
        sr = collBlank.Item(i) + 1
        er = collBlank.Item(i + 1) - 1
        CD = .Cells(sr, 11)
        Feature = .Cells(sr, 12)
        
        ptnum = er - sr + 1
        
        If ptnum = 1 Or Feature = "X" Then
        
            GoTo pthandle
            
        ElseIf Feature = "A" Then
        
            For r = sr To er
            
                cpt(0) = .Cells(r, 8)
                cpt(1) = .Cells(r, 9)
                
                Set circleobj = CAD.AddCircle(cpt, Radius)
                circleobj.Layer = "平面圖-注意點"
                
            Next
        
        Else
        
            Set lay = CAD.acadDoc.Layers.Add("平面圖-" & CD)
            
            ReDim vertices(ptnum * 3 - 1)
        
            k = 0
        
            For r = sr To er
            
                vertices(0 + 3 * k) = .Cells(r, 8)
                vertices(1 + 3 * k) = .Cells(r, 9)
                vertices(2 + 3 * k) = .Cells(r, 10)
            
                k = k + 1
                
                'If vertices(2 + 3 * k) = 0 Then
                '    Set circleobj = MyACAD.AddCircleCO(vertices(0 + 3 * k), vertices(1 + 3 * k), 0.5)
                '    circleobj.layer = "H0"
                ' End If
            
            Next
            
            Set plineobj = CAD.AddPolyLine(vertices)
            plineobj.Layer = "平面圖-" & CD
            
            Call useless.get3DCodeToSKP(vertices)
            
            'Set plineobj = MyACAD.Add3dPoly(vertices)

            
            If Feature = "R" Then plineobj.Closed = True
            
        End If
    
pthandle:
    
    Next

End With

'ACAD.SendCommand "zoom" & vbCr & "e" & vbCr

Call ChangeLayerColor

End Sub

Private Sub ChangeLayerColor() '20200728 換圖層顏色

If CAD.CADVer = "ICAD" Then Exit Sub

With sht
    
    lr = .Cells(.Rows.Count, "N").End(xlUp).Row
    
    For r = 2 To lr
        
        CD = .Cells(r, "N")
        Feature = .Cells(r, "O")
        layercolor = .Cells(r, "P")
        
        If layercolor = "" Then layercolor = 7

        If Feature = "R" Or Feature = "" Then
            CAD.acadDoc.Layers("平面圖-" & CD).Color = layercolor
        End If
        
    Next
    
    
End With

End Sub



