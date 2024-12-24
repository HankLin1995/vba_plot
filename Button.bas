Attribute VB_Name = "Button"
'20210716 HankLin整理完成
'重構程式碼真他媽的累但是可擴充性會變高
'但改好的那一刻真他媽的爽(@20210719)

'=============橫斷標竿+水準儀===========
Sub TransposeMain()

Dim obj As New clsLevelDeal
obj.transposeData

End Sub

Sub CollecMain()

Dim obj As New clsLevelDeal

obj.collectData
obj.interpolateData
obj.ExtractToLSectionSort

End Sub

'==============縱斷面====================

Sub getPFform()

Profile_Form.Show

End Sub

Sub cmdGetPlanDiff()

Dim obj As New clsLongitudinal

obj.GetPlanDiff
obj.ExportToCL_deltaH

End Sub

Sub cmdGetDataFromCAD() '縱斷面CAD取值

Dim obj As New clsLongitudinalXLS

obj.RenewTable

End Sub

'============水理因素表============================

Sub cmdGetOpenChPDF()

With Sheets("水理因素表")

    For r = 3 To 7 '.Cells(.Rows.Count, 1).End(xlUp).Row
        
        Dim obj As New clsOpenCh
        
        obj.getPropertiesByRow (r)
        obj.Calc_Report
        
        Call obj.ToPDF("水理因素表", .Cells(r, 1), "水理(空白)")
        
    Next

End With

End Sub

Sub cmdGetLocAndSlope()

Dim obj As New clsOpenCh
obj.getLocAndSlope

End Sub

Sub cmdCalcEachRow()

Dim obj As New clsOpenCh

With Sheets("水理因素表")
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To 7 'lr
    
        Call obj.Calc_ActiveRow(r)
    
    Next

End With

End Sub

'============中心線(曲線相關)======================

Sub cmdGetPtFormCurve()

Dim obj As New clsCurve

obj.GetCurve
obj.CreatePoint
obj.GetPointFromCurve

End Sub

Sub cmdDimCurve()

Dim obj As New clsCurve

obj.txtheight = Val(InputBox("請輸入標註文字大小:", , 5))

obj.GetCurve
obj.GetIP
obj.DimCurve

End Sub

Sub cmdGet3PtAlignmentArc()

Dim co As Variant
Dim obj As New clsCurve
Dim CAD As New clsACAD

co = CAD.GetPoint("PT_start")
Xs = co(0): Ys = co(1)

co = CAD.GetPoint("PT_IP")
Xm = co(0): Ym = co(1)

co = CAD.GetPoint("PT_end")
Xe = co(0): Ye = co(1)

Radius = Val(CAD.GetString("Radius=?"))

Call obj.PlotAlignmentArc(Xs, Ys, Xm, Ym, Xe, Ye, Radius)

End Sub

Sub cmdCreateCurvePolyline()

Dim obj As New clsCurve

obj.GetCurve
obj.GetCL
obj.GetIP
obj.CreatePoint
obj.ChangeCLpoint
obj.ClearAll

End Sub

'============中心線(土石方計算)=====================

Sub cmdCalcCABA()

Dim obj As New clsStone

obj.CalcCABA

End Sub

Sub cmdShowCABA()

Dim obj As New clsStone

obj.DrawCABA_Main

End Sub

Sub cmdWorkEarthReport()

Dim obj As New clsStone

obj.Report_Main

End Sub

'============中心線==================================

Sub cmdGetSecOnCAD()

Dim obj As New clsHsec_CAD
Dim IsFromTable As Boolean

msg = MsgBox("請問是否要從總表中的點資料進行對應?", vbYesNo)
If msg = vbYes Then IsFromTable = True

obj.getEGs
obj.getPLs
obj.getEGTable (IsFromTable)

Dim obj2 As New clsHsec_Table
obj2.ExtractENVcode

End Sub

Sub cmdClearSta()

Dim obj As New clsCL

obj.ClearStation

End Sub

Sub cmdfrmDataShow()

FlatMap.Show

End Sub

Sub cmdSearchLoc() '測試單獨點資料取得樁號

Dim CAD As New clsACAD
Dim CLobj As New clsCL

Call CLobj.getCenterLine

myans = "Y"

Do While myans = "Y"

    pt = CAD.GetPoint("請點選樁號查詢位置")
    Call CLobj.getSpecificLoc(pt(0), pt(1))
    
    On Error GoTo CANCLEHANDLE
    
    myans = CAD.GetString("是否要繼續查樁號?(Y/N)")
    
Loop

CANCLEHANDLE:

End Sub


'============CAD展EXCEL===============================

Sub cmdXLS2CADTable()

Dim obj As New clsExcelToCAD

obj.ExportToCAD

End Sub

'============出圖模組相關================================

Sub setLayoutDetail()

Dim la As New clsLayout

la.MyStyleSheet = InputBox("Monochrome.ctb")

Call la.setLayoutDetail

End Sub

Sub cmdMSLayout()

Dim IsAutoCAD As Boolean

IsAutoCAD = Sheets("總表").optAutoCAD

Dim la As New clsLayout

la.X = 0.8
la.Y = 5
la.dx = 390
la.dy = 258.5

la.MyConfigName = "DWG to PDF.pc5"
la.MyCanonicalMediaName = "ISO_expand_A3_(297.00_x_420.00_MM)"
la.MyStyleSheet = InputBox("出圖型式表=?", , "Monochrome.ctb")

If IsAutoCAD Then la.MyConfigName = "DWG to PDF.pc3"

la.getLayoutXLS

MsgBox "請移動至CAD模型空間窗選視埠!!"

la.CollectMSLayout
la.ClearVport
la.SortLayout (IsAutoCAD) 'ZWCAD最後排序比較方式要改成<
la.FillInLayout

End Sub

Sub cmdMSVport()

Dim la As New clsLayout

la.dx = 390
la.dy = 258.5
la.CreateMSVport

End Sub
'===========簡易橫斷面===========

Sub cmdSHTableExport_left2right()

Dim obj As New clsSimpleHSection

obj.myWidth = InputBox("請輸入中心線所跨越的寬度(公尺)")
obj.ClearHSection

For r = 4 To Cells(Rows.Count, 1).End(xlUp).Row Step 4

    obj.TableRow = r
    obj.ReadSHTable
    obj.SHTable2Hsection
    
Next

End Sub

Sub cmdSHTableExport()

Dim objSH As New clsSHtmp
Dim objH As New clsHSection

objH.ClearHSection
objSH.myWidth = Val(InputBox("請輸入中心線所跨越的寬度(公尺)"))
objSH.Export
objSH.ChangeLoc

End Sub

Sub cmdSHTableSHow()

SHTable.Show

End Sub

'============橫斷面==============

Sub cmdSHAddPoint() '樁號展點

Dim obj As New clsSimpleHSection

obj.CollectHsec
obj.GetCADLoc

End Sub

Sub cmdExtractCD() '擷取CD

Dim obj As New clsLongitudinalXLS

obj.ExtractToLSection
obj.ExtractToLSectionSort

Call cmdPlotQuickProfile

End Sub

Sub cmdPlotHsec() '繪製橫斷面

Dim obj As New clsHsec_Table
Dim CAD As New clsACAD
Dim collENV As New Collection

PaperScale = Val(InputBox("請輸入圖紙比例", , 100))

xrange = Val(InputBox("請輸入各樁號X偏移距離", , 10000))
yrange = Val(InputBox("請輸入各樁號Y偏移距離", , 2000))
times = Val(InputBox("請輸入切換個數", , 100))

Call checkBlockExist("LeftGL")
Call checkBlockExist("RightGL")
Call checkBlockExist("H_TITLE")

Set collENV = obj.getENVcoll  '取得環境資訊

BaseHeightPoint = CAD.GetPoint("Select the BaseHeightPoint")

Call obj.BatchDrawEGLine(BaseHeightPoint, PaperScale, xrange, yrange, times, collENV)

End Sub

Sub cmdPlotHsec_3D() '繪製橫斷面

Dim obj As New clsHsec_Table
Dim CAD As New clsACAD
Dim collENV As New Collection

PaperScale = 100 ' Val(InputBox("請輸入圖紙比例", , 100))

xrange = 0 'Val(InputBox("請輸入各樁號X偏移距離", , 10000))
yrange = Val(InputBox("請輸入各樁號Y偏移距離", , 200))
times = 1000 'Val(InputBox("請輸入切換個數", , 3))

Set collENV = obj.getENVcoll  '取得環境資訊

BaseHeightPoint = CAD.GetPoint("Select the BaseHeightPoint")

Call obj.BatchDrawEGLine_3D(BaseHeightPoint, PaperScale, xrange, yrange, times, collENV)

End Sub

Sub cmdPlotQuickProfile()

Dim PFobj As New clsProfile
Dim CAD As New clsACAD

PFobj.Xscale = CDbl(InputBox("請輸入X軸比例=1:", , 2500))
PFobj.Yscale = CDbl(InputBox("請輸入Y軸比例=1:", , 100))
MsgBox "請移動至CAD選取要匯入的點位"
Xstep = InputBox("請輸入樁號取樣間距", , 1)
PF_point = CAD.GetPoint("請選擇快速縱斷面基準座標")

PFobj.setProperties (PF_point)
PFobj.getHeights
PFobj.DrawHeightBar
PFobj.DrawXBar (Xstep)

End Sub


Sub cmdDefineHeight() '高程定義

Dim obj As New clsHsec_Table

obj.DefineHeight

End Sub
 
Sub cmdExtractEnvs() '匯出環境資訊 可以省略

Dim obj As New clsHsec_Table
obj.ExtractENVcode

End Sub

'============特徵線==============

Sub cmdCheckZfromTableByPL() '20210807檢查線段上是否存在沒有高程的點

Dim PLobj As New clsPL
Dim CAD As New clsACAD

r = CDbl(InputBox("請輸入注意點半徑=?", , 0.5))

Set sset = CAD.CreateSSET

For Each it In sset

    Debug.Print TypeName(it)

    If TypeName(it) = "IAcadPolyline" Or TypeName(it) Like "*LWPolyline" Then
    
        Call PLobj.getPropertiesByPL(it)
        Call PLobj.checkZFromTable(r)

    End If

Next

MsgBox "Complete!", vbInformation

End Sub

Sub cmdLineToExcel() '匯入線

Dim obj As New clsPlanMap

obj.ExportPLToExcel

End Sub

Sub cmdExcelToLine() '匯出線

Dim PL As New clsPL

With Sheets("特徵線")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        PL.createPLByRow (r)
    
    Next

End With

End Sub

'==============FIX_POINT============

Sub cmdGetPLPTByFixTable() '20210807

Dim PLobj As New clsPL
Dim CAD As New clsACAD

Set sset = CAD.CreateSSET

For Each it In sset

    Call PLobj.getPropertiesByPL(it)
    Call PLobj.createPTbyPLfromTable
    
Next

End Sub

Sub cmdPtToExcel_TYLin() '20210802

Dim obj As New clsPlanMap

obj.ExportDataToExcel_TYLin

End Sub

'============總表=================

Sub cmdPtToExcel() '匯入點位Extract point data to xls

Dim obj As New clsPlanMap

mode = InputBox("請輸入匯入模式" & vbNewLine & "1.圖塊" & vbNewLine & "2.點、文字", , 1)

Select Case mode

Case "1"
obj.ExportDataToExcel
Case "2"
obj.ExportDataToExcel_OldPT

Case "3"
obj.ExportDataToExcel_Doris

End Select

End Sub

Sub cmdDealPt() '整理數據Deal point data by their CD code

Dim obj As New clsPlanMap

mode = InputBox("請輸入排序依據：" & vbCrLf & "1.E" & vbCrLf & "2.N" & vbCrLf & "3.PT_NUM")

obj.ReArrangeCD (mode)

End Sub

Sub cmdExceltoPt() '展點Draw point data to CAD

Dim obj As New clsPlanMap

obj.ImportDataToCAD

End Sub

Sub cmdCreatePt() '補點

Dim obj As New clsPlanMap

obj.CreatePointByUser

End Sub

Sub cmdPtToLine() '連線plot line and specific feature

Dim obj As New clsPlanMap

obj.DrawLine (1)

MsgBox "連線完成!", vbInformation

End Sub


Sub cmdCADPtToExcelPt() '整理點位CAD point data to Excel point data

Dim obj As New clsPlanMap

obj.DealExportData

End Sub

Sub cmdSetDefaultFeature() 'set預設CD

Dim obj As New clsPlanMap

obj.SetDefaultFeature

End Sub

Sub cmdChangePt() '點位旋轉or平移

Dim obj As New clsPTs_Table 'clsPlanMap

mode = InputBox("請選擇移動方式" & vbNewLine & "1.平移" & vbNewLine & "2.旋轉" & vbNewLine & "3.對齊")

If mode = 1 Then
    obj.MovePoint
ElseIf mode = 2 Then
    obj.RotatePoint
ElseIf mode = 3 Then
    obj.AlignPoint
End If

End Sub

'Sub cmdImportTXT_useless()
'
''Dim obj As New clsPTs_Table 'clsPlanMap
'
''obj.ImportTXT
'
'End Sub

Sub cmdImportTXT()

    Dim filePath As String
    Dim FileContent As String
    Dim Lines() As String
    Dim Line As String
    Dim myfunc As New clsMyfunction
    
    Set sht = Sheets("總表")
    
    Call myfunc.ClearData(sht, 2, 1, 5)
    
    'FilePath = "G:\我的雲端硬碟\CADVBA\平面圖課程資料\20190328.asc"
    
    If filePath = "" Then filePath = Application.GetOpenFilename
    
    If filePath = "False" Then MsgBox "未選擇檔案!", vbCritical: End

    Open filePath For Input As #1
    Do Until EOF(1)
        Line Input #1, Line
        FileContent = FileContent & Line & vbCrLf
    Loop
    Close #1

    Lines = Split(FileContent, vbCrLf)
    
    prompt = Lines(0)
    
    mode = InputBox(prompt & vbNewLine & vbNewLine & "請選擇模式:" & vbNewLine & "1--P,E,N,Z,CD" & vbNewLine & "2--P,N,E,Z,CD", , "1")
    
    If mode = "1" Then
    
        For Each L In Lines
        
            tmp = Split(L, ",")
            
            If L <> "" Then Call myFunc.AppendData(sht, tmp)
            
        Next
    
    Else

        For Each L In Lines
        
            tmp = Split(L, ",")
            
            If UBound(tmp) > 2 Then
            
                tmp_s = tmp(1)
                tmp(1) = tmp(2)
                tmp(2) = tmp_s
                
                Call myFunc.AppendData(sht, tmp)
            
            End If


        Next
    
    End If
    
End Sub

Sub cmdExportCSV()

Dim obj As New clsPTs_Table 'clsPlanMap

obj.ExportCSV

End Sub

Sub cmdInterPolatePL()

Dim CAD As New clsACAD
Dim obj As New clsPL 'clsPlanMap

Set PL = CAD.CreateSSET("SS1")(0)

Call obj.getPropertiesByPL(PL)

mode = CAD.GetString("請選擇模式" & vbNewLine & "1.部分" & vbNewLine & "2.整段")

If mode = 1 Then
    obj.interpolatePLCoorMix
ElseIf mode = 2 Then
    obj.interpolatePLCoor
End If

End Sub
