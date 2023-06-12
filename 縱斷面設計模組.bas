Attribute VB_Name = "�a�_���]�p�Ҳ�"

Sub GetSlope()

Dim loc As Double, GlobalH As Double, LocNext As Variant, GlobalHNext As Double

IsEnd = False '�P�_�O�_����
IsDesign = False '�P�_�O�_�~��]�p
IsClear = MsgBox("�ݭn�M�ŭp�e�����ܡH" & vbCrLf & "�Y�_�A�h�|�̾ڤw�s�b���p�e�������̾A�Y��", vbYesNo, "Excel�a�_���򥻸��")
Numdigit = 2
sr = 2

With ActiveSheet

    lr = .Cells(100, sr).End(xlUp).Row
    Lc = .Cells(sr, 1).End(xlToRight).Column
    lc_check = .Cells(sr, 200).End(xlToLeft).Column

    If Not ActiveSheet.Name Like "�a�_��*" Then
        MsgBox "�u�@����~", vbCritical
        Exit Sub
    ElseIf Lc <> lc_check Then Exit Sub
        MsgBox "�Ӥu�@��Ĥ@�C���榳�ť����", vbCritical
        Exit Sub
    End If

    For r = 1 To .Cells(100, 1).End(xlUp).Row '���o�Y���C�B�]�p�_�l�I�C
    
        Select Case .Cells(r, 1)
        
            Case "�Y��": rSlope = r
            Case "�]�p�_�l�I": rShow = r
            Case "�a�L��": rGlobal = r
            Case "�p�e��": rDesign = r
            'Case "�������t": rDrop = r '���ᦳ�ŦA�ӥ�
        
        End Select
    
    Next
    
    If rSlope = "" Then rSlope = InputBox("�п�J�Y�����")
    If rShow = "" Then rShow = InputBox("�п�J�]�p�_�l�I���")
    If rGlobal = "" Then rGlobal = InputBox("�п�J�a�L�����")
    If rDesign = "" Then rDesign = InputBox("�п�J�p�e�����")
    
    If IsClear = vbYes Then .Cells(rDesign, 2).Resize(3, Lc) = "" '�����M�Ū��ʧ@
    
    For c = 2 To Lc
    
        loc = TranLoc(.Cells(sr, c))
        GlobalH = .Cells(rGlobal, c)
        DesignH = .Cells(rDesign, c)
        DesignSE = .Cells(rShow, c)
        .Cells(sr + 1, c) = 0 '��Z
        If c > 2 Then .Cells(sr + 1, c) = TranLoc(.Cells(sr, c)) - TranLoc(.Cells(sr, c - 1)) '��Z

        Select Case DesignSE
        
            Case "S", "C"
                
                If DesignH = "" Then: DesignH = InputBox("���I:" & Format(loc, "0+000") & ",�a�L��:" & GlobalH & vbCrLf & "�п�J�w�w�p�e��:", "Excel�a�_���򥻸��")
                
                IsDesign = True
                cNext = .Cells(rShow, c).End(xlToRight).Column
                GlobalHNext = .Cells(rGlobal, cNext)
                DesignHNext = .Cells(rDesign, cNext)
                LocNext = .Cells(sr, cNext)
                
                If .Cells(sr, cNext) Like "*+*" Then LocNext = TranLoc(.Cells(sr, cNext))
                
                If DesignHNext = "" Then
                
                    FitSlope = Int(-(loc - LocNext) / (Val(DesignH) - GlobalHNext))
                    
                    prompt = "���I:" & Format(loc, "0+000") & ",�a�L��:" & GlobalH & ",�p�e��:" & DesignH & vbCrLf & _
                             "���I:" & Format(LocNext, "0+000") & ",�a�L��:" & GlobalHNext & vbCrLf & _
                             "*******************" & vbCrLf & _
                             "�̾A�Y�׬�:" & FitSlope
                
                Else
                
                    FitSlope = Int(-(loc - LocNext) / (Val(DesignH) - DesignHNext))
                    
                    prompt = "���I:" & Format(loc, "0+000") & ",�a�L��:" & GlobalH & ",�p�e��:" & DesignH & vbCrLf & _
                             "���I:" & Format(LocNext, "0+000") & ",�a�L��:" & GlobalHNext & ",�p�e��:" & DesignHNext & vbCrLf & _
                             "*******************" & vbCrLf & _
                             "�̾A�Y�׬�:" & FitSlope
                
                End If
                
                slope = InputBox(prompt & vbCrLf & "�п�J�A�һݭn���Y��", "Excel�a�_���򥻸��")
                .Cells(rSlope, c) = slope

            Case "E"
            
                cNext = .Cells(rShow, c).End(xlToRight).Column
                If cNext > Lc Then IsEnd = True
                IsDesign = False
                .Cells(rSlope, c) = "End"
                
        End Select
        
        If DesignH <> "" Then
        
            .Cells(rDesign, c) = DesignH '�p�e��
            .Cells(sr + 4, c) = Round(DesignH - GlobalH, Numdigit) '�����
        
        End If
        
        If IsEnd = True Then Exit For
        If IsDesign = True Then .Cells(rDesign, c + 1) = Round(DesignH - .Cells(sr + 1, c + 1) / slope, Numdigit)
        
    Next
    
    Call ChangeLoc(sr, Lc)

End With

End Sub

Sub ChangeLoc(ByVal sr As Integer, ByVal Lc As Integer)

With ActiveSheet

    For c = Lc To 2 Step -1
        
        .Cells(sr + 1, c) = .Cells(sr + 1, c)
        
        FrontLoc = .Cells(sr, c).Value
        BackLoc = .Cells(sr, c - 1).Value
        
        d = Format(FrontLoc, "0+000.0")
        
        If FrontLoc = BackLoc Then
        
            .Cells(sr, c) = d & "(�U)"
            IsChanged = True
            
        ElseIf IsChanged = True Then
            
            .Cells(sr, c) = d & "(�W)"
            IsChanged = False
            
        Else
        
            .Cells(sr, c) = d
            
        End If
        
    Next

End With

End Sub

Sub DrawLongitudinal()

Dim Lsec As New clsLongitudinal
Dim ret(2) As Double

With Lsec 'unit mm

    .txtheight = 3 '��r�j�p
    .Interval = 2 '���j
    .VHeight = 16 '���j�p
    .startInterval = 5 '�}�l���j
    .TitleWidth = 22 '���Y�e��
    .TableMaxHeight = 277
     
    .Xscale = Val(InputBox("�а�X�b��Ҭ�" & vbCrLf & "1:", , 2500))
    .Yscale = Val(InputBox("�а�Y�b��Ҭ�" & vbCrLf & "1:", , 50))
    
    .sc = Val(InputBox("�п�J�}�l���", , 2)) '�}�l���
    .Lc = Val(InputBox("�п�J�������", , Cells(2, 1).End(xlToRight).Column))
    
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

End Sub




