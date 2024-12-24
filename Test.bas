Attribute VB_Name = "Test"
Sub cmdRebuildPLs()

Dim CAD As New clsACAD

Set sset = CAD.CreateSSET("SS1")

For Each PL In sset

    Dim PLobj As New clsPL
    
    Call PLobj.getPropertiesByPL(PL)
    Call PLobj.rebuildPL

Next

End Sub


Sub checkBlockExist(blockName)

'需要先將D:\CADWord\BLOCK資料夾中放置"Point.dwg"

Dim CAD As New clsACAD

blockFound = False
    
For Each Block In CAD.acadDoc.Blocks

    If Block.Name = blockName Then
        blockFound = True
        Exit For
    End If
    
Next Block

If Not blockFound Then
    
    MsgBox "未找到圖面具有【" & blockName & "】圖塊!嘗試創立中...", vbInformation
    
    insertPoint = Array(0, 0, 0)
    blockFilePath = "D:\CADWork\BLOCK\" & blockName & ".dwg"

    If Dir(blockFilePath) = "" Then
    
        blockFilePath = Application.GetSaveAsFilename(, , , "請選擇" & blockName & ".dwg")
        If blockFilePath = False Then MsgBox "缺少圖塊【" & blockName & ".dwg】...無法運行!", vbCritical: End
    
    End If
    
    On Error GoTo ERRORHANDLE '即便有錯誤但還是會生出Point?
    Call CAD.acadDoc.ModelSpace.InsertBlock(CAD.tranPoint(insertPoint), blockFilePath, 1, 1, 1, 0)
    
End If

ERRORHANDLE:

End Sub

Sub test_setLayout()

Dim CAD As New clsACAD

Set layout = CAD.acadDoc.Layouts

End Sub

Sub test_getPLs_Addpoint() '20221227

Dim CAD As New clsACAD
Set PLs = CAD.CreateSSET("SS1")

For Each PL In PLs

    Dim obj As New clsPL
    Call obj.getPropertiesByPL(PL)
    Call obj.addPointByPL

Next

End Sub

Sub test_createXY() '20221227

Dim o As New clsPt
Dim CAD As New clsACAD

cnt = InputBox("要放樣幾次")

If cnt = "" Then Exit Sub

For i = 1 To cnt

pt = CAD.GetPoint("請點選要標註的點資料")

x = Round(pt(0), 3)
y = Round(pt(1), 3)

Call o.getPropertiesByUser("放樣", x, y, x, y)
Call o.CreatePoint(0.5)
Call o.AppendData

Next

End Sub

Sub test_createXYpt(ByVal x As Double, ByVal y As Double)

Dim CAD As New clsACAD

Call CAD.AddPointCO(x, y)

Dim txtpt(2) As Double

txtpt(0) = x + 10
txtpt(1) = y - 5

Call CAD.AddText("X=" & CStr(x), txtpt, 5)

txtpt(0) = x + 10
txtpt(1) = y - 15

Call CAD.AddText("Y=" & CStr(y), txtpt, 5)

End Sub


Sub test_TIN_points()

Dim CAD As New clsACAD

With Sheets("總表")

Dim pt(2) As Double

For r = 2 To 28

    pt(0) = .Cells(r, 2)
    pt(1) = .Cells(r, 3)
    pt(2) = .Cells(r, 4)

    Set ptObj = CAD.AddPoint(pt)

Next

End With

End Sub

Sub test_findName()

'Debug.Print Names("UserName2").Value

Dim o As New clsFetchURL

myMac = o.getMacAddress

Names("UserName").Value = myMac

Debug.Print o.checkIsExist(myMac)

End Sub


Sub test_createMSLayout()

Dim obj As New clsLayout2
obj.getMSVport

End Sub

Sub test_getMSLayout()

Dim la As New clsLayout2

la.MyConfigName = "DWG to PDF.pc3"
la.MyCanonicalMediaName = "ISO_expand_A3_(297.00_x_420.00_MM)"
la.MyStyleSheet = InputBox("出圖型式表=?", , "Monochrome.ctb")

'obj.getMSVport
la.ClearVport
la.SortLayout True
la.FillInLayout

End Sub


Sub test_GetPLxEG()

Dim obj As New clsHsec_CAD
Dim IsFromTable As Boolean

msg = MsgBox("請問是否要從總表中的點資料進行對應?", vbYesNo)
If msg = vbYes Then IsFromTable = True

obj.getEGs
obj.getPLs
obj.getEGTable (IsFromTable)

End Sub

Sub test_PlotEGs()

Dim CAD As New clsACAD
Dim obj As New clsHsec_Table

'obj.BatchDrawEGLine

End Sub

Sub test_PlotEG()

Dim obj As New clsHsec
Dim CAD As New clsACAD
'BaseHeightPoint = CAD.GetPoint("Select the BaseHeightPoint")

Dim basept(2) As Double

x = 0 'BaseHeightPoint(0)
y = 0 ' BaseHeightPoint(1)

Call obj.setBaseHeightPT(basept)  'CAD.GetPoint("Select the BaseHeightPoint"))
Call obj.getPropertiesByRows(2, 13)
Call obj.plotHsec
Call obj.plotOther(collENV)
Call obj.DrawHeightBar
Call obj.plotTitle

End Sub

Sub test_rebuildPLs()

Dim CAD As New clsACAD
Dim obj As New clsPL

Call obj.createPLByRow(2)

End Sub

Sub test_getPLs()

Dim CAD As New clsACAD
Set PLs = CAD.CreateSSET("SS1")

For Each PL In PLs

    Dim obj As New clsPL
    Call obj.getPropertiesByPL(PL)
    Call obj.AppendData

Next

End Sub

Sub test_addlayer()

Dim CAD As New clsACAD

With Sheets("SET")

    lr = .Cells(.Rows.Count, 9).End(xlUp).Row
    
    For r = 2 To 26 'lr
    
        'Debug.Print .Cells(r, "L")
    
        Set lay = CAD.acadDoc.Layers.Add(.Cells(r, 9))

        Select Case .Cells(r, 9 + 2)

        Case "紅": layercolor = 1
        Case "黃": layercolor = 1 + 1
        Case "綠": layercolor = 1 + 2
        Case "青": layercolor = 1 + 3
        Case "藍": layercolor = 1 + 4
        Case "粉紅": layercolor = 1 + 5
        Case "白": layercolor = 1 + 6
        Case "灰": layercolor = 253
        Case "中心紅": layercolor = 10

        End Select

        lay.Color = layercolor

        If .Cells(r, 9 + 3) <> "" Then

            lay.linetype = .Cells(r, 9 + 3)

        End If
        
    Next

End With

End Sub

Sub test_checkPointElev() '20210615

Dim CAD As New clsACAD
Dim pmObj As New clsPlanMap

Set o = CAD.CreateSSET()

rr = 2

For Each PL In o

    pts = PL.Coordinates
    co = 3: If TypeName(PL) Like "*LWPolyline" Then co = 2
    
    For i = 0 To UBound(pts) Step co

        X0 = pts(i): Y0 = pts(i + 1)
        Z = pmObj.CollPointTable(X0 & ":" & Y0)
        
        With Sheets("TMP")
            .Cells(rr, 1) = rr - 1
            .Cells(rr, 2) = X0
            .Cells(rr, 3) = Y0
            .Cells(rr, 4) = Z
            .Cells(rr, 5) = "PL" & k
        
        End With
        
        If Z = 0 Then
            Set circleobj = CAD.AddCircleCO(X0, Y0, 0.5)
            circleobj.Layer = "平面圖-注意點"
        End If

        rr = rr + 1

    Next
    
    k = k + 1
    
Next

End Sub

Sub test_CLPointByUser()

'With Me

'Me.Hide

Dim obj As New clsCL

'If Not .chkXLS Then obj.nowLoc = .txtStartLoc
obj.nowLoc = InputBox("請輸入起點樁號", , 0) ' .txtStartLoc
obj.w = 1 ' .txtWdith
'obj.IsLeftShow = .chkLeftShow
'obj.IsRightShow = .chkRightShow
'obj.NeedBox = .chkNeedBox
'obj.NeedDir = .chkNeedDir
'obj.NeedReverse = .chkReverse
obj.PaperScale = 1000 '.tboPaperScale
obj.WIDTH_COE = 1.2 '.tboWidthCoe

obj.getCenterLine

'If Not .chkXLS Then
    obj.getLoc
'Else
   ' obj.getLocXLS
'End If

obj.CrossLine_Main_Output
'obj.setDataByUser

'End With

'Unload Me

End Sub

Sub test_getPoint()

Dim CAD As New clsACAD
Dim pt As New clsPt

Set sset = CAD.CreateSSET("SS1")

For Each it In sset

    Call pt.getPropertiesByBlock_Tien(it) ', PT_NUM, E, N)
    
    pt.AppendData

Next

End Sub

Sub test_getAreas()

Dim CAD As New clsACAD
Dim myMath As New clsMath

Set sset = CAD.CreateSSET


For Each it In sset
    
    j = j + 1

    o = myMath.getCentroid(it)
    
    'Set pointObject = CAD.AddPoint(o)
    
    x = o(0)
    y = o(1)
    
    Dim txtpt(2) As Double
    
    txtpt(0) = x
    txtpt(1) = y - 200
    
    BoundaryArea = Round(it.area / 10000, 2) '& "m2"

    Set txtobject = CAD.AddText(BoundaryArea, txtpt, 200, 2)
    Call CAD.AddTextBox(txtobject)
    
    txtpt(0) = x
    txtpt(1) = y + 200
    
    Call CAD.AddText("A" & j, txtpt, 200, 2)

Next

End Sub

Sub test_getHandles()

Dim CAD As New clsACAD
Dim myMath As New clsMath

Set sset = CAD.CreateSSET

For Each it In sset
    
    j = j + 1

    o = myMath.getCentroid(it)
    
    oo = it.Handle
    
    Dim txtpt(2) As Double
    
    Set txtobject = CAD.AddText(oo, o, 200, 2)
    
    With Sheets("AREA")
    
        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        .Cells(lr + 1, 1) = "'" & oo
        .Cells(lr + 1, 2) = o(0)
        .Cells(lr + 1, 3) = o(1)
    
    End With
    
Next

End Sub

Sub test_AboutHandle()

Dim CAD As New clsACAD

Dim Math As New clsMath

With Sheets("AREA")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lr

        handleInput = .Cells(r, 1)
        
        Set acadDoc = CAD.acadDoc
        Set objByHandle = acadDoc.HandleToObject(handleInput)
        
        o = Math.getCentroid(objByHandle)
        
        BoundaryArea = Round(objByHandle.area / 10000, 2) '& "m2"
        
        .Cells(r, 5) = BoundaryArea

        'Dim txtpt(2) As Double
        
        'txtpt(0) = o(0)
        'txtpt(1) = o(1) - 200
        
        'Set txtobj = CAD.AddText(r - 1, txtpt, 200, 2)

    Next

End With


End Sub

