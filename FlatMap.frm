VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FlatMap 
   Caption         =   "平面圖工具"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5010
   OleObjectBlob   =   "FlatMap.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "FlatMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub chkXLS_Click()

If Me.chkXLS = True Then
    Me.txtStartLoc = ""
    Me.txtStartLoc.Enabled = False
Else
    Me.txtStartLoc.Enabled = True
End If

End Sub

Private Sub cmdCreateCenterLine_Click()

Me.Hide

Dim obj As New clsCL

obj.BLnext = Me.txtCenterDis
obj.BorderLine_Main
obj.DrawCenterLine

Me.Show

End Sub

Private Sub cmdGetLoc_Click()

With Me

Me.Hide

Dim obj As New clsCL

If Not .chkXLS Then obj.nowLoc = .txtStartLoc
obj.w = .txtWdith
obj.IsLeftShow = .chkLeftShow
obj.IsRightShow = .chkRightShow
obj.NeedBox = .chkNeedBox
obj.NeedDir = .chkNeedDir
obj.NeedReverse = .chkReverse
obj.PaperScale = .tboPaperScale
obj.WIDTH_COE = .tboWidthCoe

obj.getCenterLine

If Not .chkXLS Then
    obj.getLoc
Else
    obj.getLocXLS
End If

obj.CrossLine_Main
obj.setDataByUser

End With

Unload Me

End Sub

Sub oldcode()

'Me.Hide

'Dim obj2 As New clsLocationLine

'obj.MapScale = 1000 'mm2m
'obj.StartLoc = Me.txtStartLoc
'obj.width = Me.txtWdith
'obj.txtheight = Me.txtTxtHeight

'obj.SelectLine
'obj.CreateCrossLine

'Me.Show

'Dim obj As New clsLocationLine

'obj.MapScale = 1000 'mm2m
'obj.width = 100
'obj.NextDistance = Val(Me.txtCenterDis)
'obj.SelectBorderLine
'obj.CreateBorderLine
'obj.DrawCenterLine

End Sub

