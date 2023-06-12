Attribute VB_Name = "縱斷面設計模組"

Sub GetSlope()

Dim loc As Double, GlobalH As Double, LocNext As Variant, GlobalHNext As Double

IsEnd = False '判斷是否結束
IsDesign = False '判斷是否繼續設計
IsClear = MsgBox("需要清空計畫高欄位嗎？" & vbCrLf & "若否，則會依據已存在的計畫高給予最適坡度", vbYesNo, "Excel縱斷面基本資料")
Numdigit = 2
sr = 2

With ActiveSheet

    lr = .Cells(100, sr).End(xlUp).Row
    Lc = .Cells(sr, 1).End(xlToRight).Column
    lc_check = .Cells(sr, 200).End(xlToLeft).Column

    If Not ActiveSheet.Name Like "縱斷面*" Then
        MsgBox "工作表錯誤", vbCritical
        Exit Sub
    ElseIf Lc <> lc_check Then Exit Sub
        MsgBox "該工作表第一列不行有空白欄位", vbCritical
        Exit Sub
    End If

    For r = 1 To .Cells(100, 1).End(xlUp).Row '取得坡降列、設計起始點列
    
        Select Case .Cells(r, 1)
        
            Case "坡降": rSlope = r
            Case "設計起始點": rShow = r
            Case "地盤高": rGlobal = r
            Case "計畫高": rDesign = r
            'Case "直接落差": rDrop = r '之後有空再來用
        
        End Select
    
    Next
    
    If rSlope = "" Then rSlope = InputBox("請輸入坡降欄位")
    If rShow = "" Then rShow = InputBox("請輸入設計起始點欄位")
    If rGlobal = "" Then rGlobal = InputBox("請輸入地盤高欄位")
    If rDesign = "" Then rDesign = InputBox("請輸入計畫高欄位")
    
    If IsClear = vbYes Then .Cells(rDesign, 2).Resize(3, Lc) = "" '首次清空的動作
    
    For c = 2 To Lc
    
        loc = TranLoc(.Cells(sr, c))
        GlobalH = .Cells(rGlobal, c)
        DesignH = .Cells(rDesign, c)
        DesignSE = .Cells(rShow, c)
        .Cells(sr + 1, c) = 0 '單距
        If c > 2 Then .Cells(sr + 1, c) = TranLoc(.Cells(sr, c)) - TranLoc(.Cells(sr, c - 1)) '單距

        Select Case DesignSE
        
            Case "S", "C"
                
                If DesignH = "" Then: DesignH = InputBox("測點:" & Format(loc, "0+000") & ",地盤高:" & GlobalH & vbCrLf & "請輸入預定計畫高:", "Excel縱斷面基本資料")
                
                IsDesign = True
                cNext = .Cells(rShow, c).End(xlToRight).Column
                GlobalHNext = .Cells(rGlobal, cNext)
                DesignHNext = .Cells(rDesign, cNext)
                LocNext = .Cells(sr, cNext)
                
                If .Cells(sr, cNext) Like "*+*" Then LocNext = TranLoc(.Cells(sr, cNext))
                
                If DesignHNext = "" Then
                
                    FitSlope = Int(-(loc - LocNext) / (Val(DesignH) - GlobalHNext))
                    
                    prompt = "測點:" & Format(loc, "0+000") & ",地盤高:" & GlobalH & ",計畫高:" & DesignH & vbCrLf & _
                             "測點:" & Format(LocNext, "0+000") & ",地盤高:" & GlobalHNext & vbCrLf & _
                             "*******************" & vbCrLf & _
                             "最適坡度為:" & FitSlope
                
                Else
                
                    FitSlope = Int(-(loc - LocNext) / (Val(DesignH) - DesignHNext))
                    
                    prompt = "測點:" & Format(loc, "0+000") & ",地盤高:" & GlobalH & ",計畫高:" & DesignH & vbCrLf & _
                             "測點:" & Format(LocNext, "0+000") & ",地盤高:" & GlobalHNext & ",計畫高:" & DesignHNext & vbCrLf & _
                             "*******************" & vbCrLf & _
                             "最適坡度為:" & FitSlope
                
                End If
                
                slope = InputBox(prompt & vbCrLf & "請輸入你所需要的坡度", "Excel縱斷面基本資料")
                .Cells(rSlope, c) = slope

            Case "E"
            
                cNext = .Cells(rShow, c).End(xlToRight).Column
                If cNext > Lc Then IsEnd = True
                IsDesign = False
                .Cells(rSlope, c) = "End"
                
        End Select
        
        If DesignH <> "" Then
        
            .Cells(rDesign, c) = DesignH '計畫高
            .Cells(sr + 4, c) = Round(DesignH - GlobalH, Numdigit) '挖填方
        
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
        
            .Cells(sr, c) = d & "(下)"
            IsChanged = True
            
        ElseIf IsChanged = True Then
            
            .Cells(sr, c) = d & "(上)"
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

    .txtheight = 3 '文字大小
    .Interval = 2 '間隔
    .VHeight = 16 '表格大小
    .startInterval = 5 '開始間隔
    .TitleWidth = 22 '表頭寬度
    .TableMaxHeight = 277
     
    .Xscale = Val(InputBox("請問X軸比例為" & vbCrLf & "1:", , 2500))
    .Yscale = Val(InputBox("請問Y軸比例為" & vbCrLf & "1:", , 50))
    
    .sc = Val(InputBox("請輸入開始欄位", , 2)) '開始欄位
    .Lc = Val(InputBox("請輸入結束欄位", , Cells(2, 1).End(xlToRight).Column))
    
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

End Sub




