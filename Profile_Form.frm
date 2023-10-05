VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Profile_Form 
   Caption         =   "縱斷面繪圖參數"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   OleObjectBlob   =   "Profile_Form.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Profile_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdOK_Click()

Dim Lsec As New clsLongitudinal
Dim ret(2) As Double

With Lsec 'unit mm

    .txtheight = Me.tboText  '文字大小
    .Interval = Me.tboInterval '間隔
    .VHeight = Me.tboVheight  '表格大小
    .startInterval = Me.tboStartInterval  '開始間隔
    .TitleWidth = Me.tboStartWidth  '表頭寬度
    .TableMaxHeight = 277
     
    .Xscale = Me.tboXScale 'Val(InputBox("請問X軸比例為" & vbCrLf & "1:", , 2500))
    .Yscale = Me.tboYScale ' Val(InputBox("請問Y軸比例為" & vbCrLf & "1:", , 50))
    
    .sc = Me.tbosc ' Val(InputBox("請輸入開始欄位", , 2)) '開始欄位
    .Lc = Me.tboec 'Val(InputBox("請輸入結束欄位", , Cells(2, 1).End(xlToRight).Column))
    
    Call ChangeLoc(2, .Lc)
    
    .ReadData
    .GetScale '(ret) '附帶有基準點的sub
    .DrawOuter
    .FillInTable
    .DrawHeightBar
    .FillInNote
    .DrawHeight
    .FillInSlopeAndSE
    .TableIntroduce

End With

Unload Me

End Sub

Private Sub UserForm_Initialize()

tboec.Text = Sheets("縱斷面繪圖").Cells(2, 1).End(xlToRight).Column

End Sub
