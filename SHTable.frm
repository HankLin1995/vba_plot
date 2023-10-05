VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SHTable 
   Caption         =   "紅白標竿橫斷紀錄"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7875
   OleObjectBlob   =   "SHTable.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "SHTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Sub cmdAdd_Click()

With Me

If .txtX = "" Or .txtY = "" Then MsgBox "X軸或Y軸不得為空白，至少需輸入0": Exit Sub

i = .lstTMP.ListCount

.lstTMP.AddItem i

.lstTMP.List(i, 0) = .txtX
.lstTMP.List(i, 1) = .txtY
.lstTMP.List(i, 2) = .txtCD

If .txtCD = "" Then .lstTMP.List(i, 2) = "NULL"

.txtX = ""
.txtY = ""
.txtCD = ""

.txtX.SetFocus

End With

End Sub

Private Sub cmdDel_Click()

i = Me.lstTMP.ListIndex

Me.lstTMP.RemoveItem (i)

End Sub

Private Sub cmdOK_Click()

Dim obj As New clsSHtmp

obj.ImportData

End Sub

Private Sub CommandButton1_Click()

Dim a As New clsSHtmp

a.Test

End Sub

Private Sub OptionButton1_Click()

Me.Image1.Visible = True

End Sub

Private Sub OptionButton2_Click()

Me.Image1.Visible = False

End Sub


