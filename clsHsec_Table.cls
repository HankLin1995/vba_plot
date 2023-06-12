VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHsec_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===2021/11/26 NOTE===
'===Author:HankLin===

'BatchDrawEGLine
'BatchDrawEGLine_3D
'ExtractENVcode >>將橫斷面資料表的左右邊界匯入至環境狀態列表
'DefineHeight >>由自訂高程代碼進行取代運算
'getENVcoll >> 確認環境狀態是否為空值

Private CAD As New clsACAD
Private myFunc As New clsMyfunction
Private sht_EG As Object

Private Sub Class_Initialize()

Set sht_EG = Sheets("橫斷面")

End Sub

Sub BatchDrawEGLine(ByVal BaseHeightPoint, ByVal PaperScale, ByVal xrange, ByVal yrange, ByVal times, ByVal collENV) '批次繪製橫斷面，為主要控制sub。

Dim CL As New clsCL
CL.getDataByRng
SectionWidth = CL.w

X0 = BaseHeightPoint(0)
Y0 = BaseHeightPoint(1)

With sht_EG

Set collBlank = myFunc.getBlankColl(sht_EG, 1)

For i = 1 To collBlank.Count - 1
    
    '=======繪製橫斷面=========
    
    sr = collBlank.Item(i) + 1
    er = collBlank.Item(i + 1) - 1
    
    Dim obj As New clsHsec
    
    obj.PaperScale = PaperScale
    obj.SectionWidth = SectionWidth
    obj.Is3D = False
    
    Call obj.setBaseHeightPT(BaseHeightPoint)
    Call obj.getPropertiesByRows(sr, er)
    Call obj.plotHsec
    Call obj.plotUchannel
    'Call obj.plotUchannel_Block
    Call obj.plotOther(collENV)
    Call obj.DrawHeightBar
    Call obj.plotTitle
    
    nowStep = obj.getStep(sr, er)
    
    '=======取得樁號間的Y距離=============
    
    If i = collBlank.Count - 1 Then Exit For
    
    Dim obj2 As New clsHsec
    
    sr = collBlank.Item(i + 1) + 1
    er = collBlank.Item(i + 2) - 1
    
    nextStep = obj2.getStep(sr, er)
    
    ydistance = (nowStep + nextStep) * 1000 + yrange
    
    'ydistance = yrange
    
    '======取得樁號間的X距離============
    
    cnt = cnt + 1

    If cnt Mod times = 0 Then

        xdistance = xdistance + 1
        BaseHeightPoint(1) = Y0 + ydistance

    End If
    
    '===================================

    BaseHeightPoint(0) = X0 + xdistance * xrange
    BaseHeightPoint(1) = BaseHeightPoint(1) - ydistance * (PaperScale / 100)
    
Next

End With

End Sub

Sub BatchDrawEGLine_3D(ByVal BaseHeightPoint, ByVal PaperScale, ByVal xrange, ByVal yrange, ByVal times, ByVal collENV)
'批次繪製橫斷面，為主要控制sub。

Dim CL As New clsCL
CL.getDataByRng
SectionWidth = CL.w

X0 = BaseHeightPoint(0)
Y0 = BaseHeightPoint(1)

With sht_EG

Set collBlank = myFunc.getBlankColl(sht_EG, 1)

For i = 1 To collBlank.Count - 1
    
    '=======繪製橫斷面=========
    
    sr = collBlank.Item(i) + 1
    er = collBlank.Item(i + 1) - 1
    
    Dim obj As New clsHsec
    
    obj.PaperScale = PaperScale
    obj.SectionWidth = SectionWidth
    obj.Is3D = True
    
    Call obj.setBaseHeightPT(BaseHeightPoint)
    Call obj.getPropertiesByRows(sr, er)
    Call obj.plotHsec
    Call obj.plotUchannel
    Call obj.plotOther(collENV)
    'Call obj.DrawHeightBar
    'Call obj.plotTitle
    
    nowStep = obj.getStep(sr, er)
    
    '=======取得樁號間的Y距離=============
    
    If i = collBlank.Count - 1 Then Exit For
    
    Dim obj2 As New clsHsec
    
    sr = collBlank.Item(i + 1) + 1
    er = collBlank.Item(i + 2) - 1
    
    nextStep = obj2.getStep(sr, er)
    
    'ydistance = (nowStep + nextStep) * 1000 + yrange
    
    ydistance = yrange
    
    '======取得樁號間的X距離============
    
    cnt = cnt + 1

    If cnt Mod times = 0 Then

        xdistance = xdistance + 1
        BaseHeightPoint(1) = Y0 + ydistance

    End If
    
    '===================================

    BaseHeightPoint(0) = X0 + xdistance * xrange
    BaseHeightPoint(1) = BaseHeightPoint(1) - ydistance * (PaperScale / 100)
    
Next

End With

End Sub

Sub ExtractENVcode()

Dim coll As New Collection

With sht_EG

Set collBlank = myFunc.getBlankColl(sht_EG, 1)

For i = 1 To collBlank.Count - 1
    
    sr = collBlank.Item(i) + 1
    er = collBlank.Item(i + 1) - 1
    
    leftEnv = CStr(.Cells(sr, 3))
    rightEnv = CStr(.Cells(er, 3))

    On Error Resume Next
    
    coll.Add leftEnv, leftEnv
    coll.Add rightEnv, rightEnv
    
    On Error GoTo 0

Next

Call myFunc.ClearData(sht_EG, 2, 14, 2)

r = 2

For Each it In coll
    
    .Cells(r, "N") = it
    
    If it Like "*RD*" Then
        .Cells(r, "O") = "路"
    ElseIf it Like "*FA*" Then
        .Cells(r, "O") = "田"
    End If
    
    r = r + 1

Next

.Activate

End With

End Sub

Sub DefineHeight() '自訂高程CD碼對照及運算

Dim collBlank As New Collection
Dim colltarget As New Collection
Dim collreplace As New Collection
Dim colltmp As New Collection
Dim arr() As Variant

With sht_EG

    lr = .Cells(.Rows.Count, 10).End(xlUp).Row

    If lr = 1 Then Exit Sub

    ReDim arr(1 To lr - 1, 1 To 4)

    For r = 2 To .Cells(Rows.Count, 10).End(xlUp).Row
    
        arr(r - 1, 1) = .Cells(r, 10)
        arr(r - 1, 2) = .Cells(r, 11)
        arr(r - 1, 4) = .Cells(r, 12)
    
        'colltarget.Add .Cells(r, 10).Value
        'collreplace.Add .Cells(r, 11).Value
    
    Next
    
    Set collBlank = myFunc.getBlankColl(sht_EG, 1)

    For i = 1 To collBlank.Count - 1
    
        sr = collBlank.Item(i) + 1
        er = collBlank.Item(i + 1) - 1
        CD = .Cells(sr, 3)
        
        For j = 1 To UBound(arr)
        
            arr(j, 3) = 0
        
        Next
        
        For r = sr To er
        
            H = .Cells(r, 2)
            CD = .Cells(r, 3)
            
            For j = 1 To UBound(arr)
            
                If CD = arr(j, 2) Then arr(j, 3) = H
            
            Next
    
        Next
    
        For r = sr To er
        
            CD = .Cells(r, 3)
            
            For j = 1 To UBound(arr)
            
                If CD = arr(j, 1) Then .Cells(r, 2) = arr(j, 3) + arr(j, 4)
                
            Next
            
        Next
    
    Next
    
End With

End Sub

Function getENVcoll()

Dim coll As New Collection

With sht_EG

    lr = .Cells(.Rows.Count, "N").End(xlUp).Row
    
    For r = 2 To lr
    
        env_CD = CStr(.Cells(r, "N"))
        env_txt = .Cells(r, "O")
        
        If env_txt = "" Then
            env_txt = InputBox("請輸入" & env_CD & "對應環境資訊", , "田")
            .Cells(r, "O") = env_txt
        End If
    
        coll.Add env_txt, env_CD
        
    Next

End With

Set getENVcoll = coll

End Function

