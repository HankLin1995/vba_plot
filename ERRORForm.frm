VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ERRORForm 
   Caption         =   "系統錯誤回報"
   ClientHeight    =   4650
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4455
   OleObjectBlob   =   "ERRORForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ERRORForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdSubmit_Click()

'=====SignDetail==========

Dim o As New clsFetchURL

user_name = Me.tboName.Text
user_company = Me.tboJob.Text
user_mail = Me.tboMail.Text
msg = Me.tboMSG.Text

If user_name = "" Then MsgBox "請輸入使用者姓名", vbCritical: Exit Sub
If user_company = "" Then MsgBox "請輸入公司名稱", vbCritical: Exit Sub
If user_mail = "" Then MsgBox "請輸入電子郵件", vbCritical: Exit Sub
If msg = "" Then MsgBox "請輸入錯誤訊息", vbCritical: Exit Sub

myURL_GAS = o.CreateURL("ERRORMSG", user_name, user_company, user_mail, msg)
o.ExecHTTP (myURL_GAS)

MsgBox "已發送成功，等候通知!", vbInformation

ThisWorkbook.Close False

Unload Me
    
End Sub


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()

ThisWorkbook.Close False

End Sub
