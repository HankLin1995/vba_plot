VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Profile_Form 
   Caption         =   "�a�_��ø�ϰѼ�"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   OleObjectBlob   =   "Profile_Form.frx":0000
   StartUpPosition =   1  '���ݵ�������
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

    .txtheight = Me.tboText  '��r�j�p
    .Interval = Me.tboInterval '���j
    .VHeight = Me.tboVheight  '���j�p
    .startInterval = Me.tboStartInterval  '�}�l���j
    .TitleWidth = Me.tboStartWidth  '���Y�e��
    .TableMaxHeight = 277
     
    .Xscale = Me.tboXScale 'Val(InputBox("�а�X�b��Ҭ�" & vbCrLf & "1:", , 2500))
    .Yscale = Me.tboYScale ' Val(InputBox("�а�Y�b��Ҭ�" & vbCrLf & "1:", , 50))
    
    .sc = Me.tbosc ' Val(InputBox("�п�J�}�l���", , 2)) '�}�l���
    .Lc = Me.tboec 'Val(InputBox("�п�J�������", , Cells(2, 1).End(xlToRight).Column))
    
    Call ChangeLoc(2, .Lc)
    
    .ReadData
    .GetScale '(ret) '���a������I��sub
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

tboec.Text = Sheets("�a�_��ø��").Cells(2, 1).End(xlToRight).Column

End Sub
