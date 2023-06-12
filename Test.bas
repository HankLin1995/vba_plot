Attribute VB_Name = "Test"
Sub test_TIN_points()

Dim CAD As New clsACAD

With Sheets("總表")

Dim pt(2) As Double

For r = 2 To 28

    pt(0) = .Cells(r, 2)
    pt(1) = .Cells(r, 3)
    pt(2) = .Cells(r, 4)

    Set ptobj = CAD.AddPoint(pt)

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

X = 0 'BaseHeightPoint(0)
Y = 0 ' BaseHeightPoint(1)

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
    
        Set lay = CAD.acaddoc.Layers.Add(.Cells(r, 9))

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

    pts = PL.coordinates
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




