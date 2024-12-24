Attribute VB_Name = "Button"
'20210716 HankLin��z����
'���c�{���X�u�L�����֦��O�i�X�R�ʷ|�ܰ�
'����n�����@��u�L�����n(@20210719)

'=============���_�Ь�+���ǻ�===========
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

'==============�a�_��====================

Sub getPFform()

Profile_Form.Show

End Sub

Sub cmdGetPlanDiff()

Dim obj As New clsLongitudinal

obj.GetPlanDiff
obj.ExportToCL_deltaH

End Sub

Sub cmdGetDataFromCAD() '�a�_��CAD����

Dim obj As New clsLongitudinalXLS

obj.RenewTable

End Sub

'============���z�]����============================

Sub cmdGetOpenChPDF()

With Sheets("���z�]����")

    For r = 3 To 7 '.Cells(.Rows.Count, 1).End(xlUp).Row
        
        Dim obj As New clsOpenCh
        
        obj.getPropertiesByRow (r)
        obj.Calc_Report
        
        Call obj.ToPDF("���z�]����", .Cells(r, 1), "���z(�ť�)")
        
    Next

End With

End Sub

Sub cmdGetLocAndSlope()

Dim obj As New clsOpenCh
obj.getLocAndSlope

End Sub

Sub cmdCalcEachRow()

Dim obj As New clsOpenCh

With Sheets("���z�]����")
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To 7 'lr
    
        Call obj.Calc_ActiveRow(r)
    
    Next

End With

End Sub

'============���߽u(���u����)======================

Sub cmdGetPtFormCurve()

Dim obj As New clsCurve

obj.GetCurve
obj.CreatePoint
obj.GetPointFromCurve

End Sub

Sub cmdDimCurve()

Dim obj As New clsCurve

obj.txtheight = Val(InputBox("�п�J�е���r�j�p:", , 5))

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

'============���߽u(�g�ۤ�p��)=====================

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

'============���߽u==================================

Sub cmdGetSecOnCAD()

Dim obj As New clsHsec_CAD
Dim IsFromTable As Boolean

msg = MsgBox("�аݬO�_�n�q�`�����I��ƶi�����?", vbYesNo)
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

Sub cmdSearchLoc() '���ճ�W�I��ƨ��o�θ�

Dim CAD As New clsACAD
Dim CLobj As New clsCL

Call CLobj.getCenterLine

myans = "Y"

Do While myans = "Y"

    pt = CAD.GetPoint("���I��θ��d�ߦ�m")
    Call CLobj.getSpecificLoc(pt(0), pt(1))
    
    On Error GoTo CANCLEHANDLE
    
    myans = CAD.GetString("�O�_�n�~��d�θ�?(Y/N)")
    
Loop

CANCLEHANDLE:

End Sub


'============CAD�iEXCEL===============================

Sub cmdXLS2CADTable()

Dim obj As New clsExcelToCAD

obj.ExportToCAD

End Sub

'============�X�ϼҲլ���================================

Sub setLayoutDetail()

Dim la As New clsLayout

la.MyStyleSheet = InputBox("Monochrome.ctb")

Call la.setLayoutDetail

End Sub

Sub cmdMSLayout()

Dim IsAutoCAD As Boolean

IsAutoCAD = Sheets("�`��").optAutoCAD

Dim la As New clsLayout

la.X = 0.8
la.Y = 5
la.dx = 390
la.dy = 258.5

la.MyConfigName = "DWG to PDF.pc5"
la.MyCanonicalMediaName = "ISO_expand_A3_(297.00_x_420.00_MM)"
la.MyStyleSheet = InputBox("�X�ϫ�����=?", , "Monochrome.ctb")

If IsAutoCAD Then la.MyConfigName = "DWG to PDF.pc3"

la.getLayoutXLS

MsgBox "�в��ʦ�CAD�ҫ��Ŷ��������!!"

la.CollectMSLayout
la.ClearVport
la.SortLayout (IsAutoCAD) 'ZWCAD�̫�ƧǤ���覡�n�令<
la.FillInLayout

End Sub

Sub cmdMSVport()

Dim la As New clsLayout

la.dx = 390
la.dy = 258.5
la.CreateMSVport

End Sub
'===========²�����_��===========

Sub cmdSHTableExport_left2right()

Dim obj As New clsSimpleHSection

obj.myWidth = InputBox("�п�J���߽u�Ҹ�V���e��(����)")
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
objSH.myWidth = Val(InputBox("�п�J���߽u�Ҹ�V���e��(����)"))
objSH.Export
objSH.ChangeLoc

End Sub

Sub cmdSHTableSHow()

SHTable.Show

End Sub

'============���_��==============

Sub cmdSHAddPoint() '�θ��i�I

Dim obj As New clsSimpleHSection

obj.CollectHsec
obj.GetCADLoc

End Sub

Sub cmdExtractCD() '�^��CD

Dim obj As New clsLongitudinalXLS

obj.ExtractToLSection
obj.ExtractToLSectionSort

Call cmdPlotQuickProfile

End Sub

Sub cmdPlotHsec() 'ø�s���_��

Dim obj As New clsHsec_Table
Dim CAD As New clsACAD
Dim collENV As New Collection

PaperScale = Val(InputBox("�п�J�ϯȤ��", , 100))

xrange = Val(InputBox("�п�J�U�θ�X�����Z��", , 10000))
yrange = Val(InputBox("�п�J�U�θ�Y�����Z��", , 2000))
times = Val(InputBox("�п�J�����Ӽ�", , 100))

Call checkBlockExist("LeftGL")
Call checkBlockExist("RightGL")
Call checkBlockExist("H_TITLE")

Set collENV = obj.getENVcoll  '���o���Ҹ�T

BaseHeightPoint = CAD.GetPoint("Select the BaseHeightPoint")

Call obj.BatchDrawEGLine(BaseHeightPoint, PaperScale, xrange, yrange, times, collENV)

End Sub

Sub cmdPlotHsec_3D() 'ø�s���_��

Dim obj As New clsHsec_Table
Dim CAD As New clsACAD
Dim collENV As New Collection

PaperScale = 100 ' Val(InputBox("�п�J�ϯȤ��", , 100))

xrange = 0 'Val(InputBox("�п�J�U�θ�X�����Z��", , 10000))
yrange = Val(InputBox("�п�J�U�θ�Y�����Z��", , 200))
times = 1000 'Val(InputBox("�п�J�����Ӽ�", , 3))

Set collENV = obj.getENVcoll  '���o���Ҹ�T

BaseHeightPoint = CAD.GetPoint("Select the BaseHeightPoint")

Call obj.BatchDrawEGLine_3D(BaseHeightPoint, PaperScale, xrange, yrange, times, collENV)

End Sub

Sub cmdPlotQuickProfile()

Dim PFobj As New clsProfile
Dim CAD As New clsACAD

PFobj.Xscale = CDbl(InputBox("�п�JX�b���=1:", , 2500))
PFobj.Yscale = CDbl(InputBox("�п�JY�b���=1:", , 100))
MsgBox "�в��ʦ�CAD����n�פJ���I��"
Xstep = InputBox("�п�J�θ����˶��Z", , 1)
PF_point = CAD.GetPoint("�п�ܧֳt�a�_����Ǯy��")

PFobj.setProperties (PF_point)
PFobj.getHeights
PFobj.DrawHeightBar
PFobj.DrawXBar (Xstep)

End Sub


Sub cmdDefineHeight() '���{�w�q

Dim obj As New clsHsec_Table

obj.DefineHeight

End Sub
 
Sub cmdExtractEnvs() '�ץX���Ҹ�T �i�H�ٲ�

Dim obj As New clsHsec_Table
obj.ExtractENVcode

End Sub

'============�S�x�u==============

Sub cmdCheckZfromTableByPL() '20210807�ˬd�u�q�W�O�_�s�b�S�����{���I

Dim PLobj As New clsPL
Dim CAD As New clsACAD

r = CDbl(InputBox("�п�J�`�N�I�b�|=?", , 0.5))

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

Sub cmdLineToExcel() '�פJ�u

Dim obj As New clsPlanMap

obj.ExportPLToExcel

End Sub

Sub cmdExcelToLine() '�ץX�u

Dim PL As New clsPL

With Sheets("�S�x�u")

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

'============�`��=================

Sub cmdPtToExcel() '�פJ�I��Extract point data to xls

Dim obj As New clsPlanMap

mode = InputBox("�п�J�פJ�Ҧ�" & vbNewLine & "1.�϶�" & vbNewLine & "2.�I�B��r", , 1)

Select Case mode

Case "1"
obj.ExportDataToExcel
Case "2"
obj.ExportDataToExcel_OldPT

Case "3"
obj.ExportDataToExcel_Doris

End Select

End Sub

Sub cmdDealPt() '��z�ƾ�Deal point data by their CD code

Dim obj As New clsPlanMap

mode = InputBox("�п�J�ƧǨ̾ڡG" & vbCrLf & "1.E" & vbCrLf & "2.N" & vbCrLf & "3.PT_NUM")

obj.ReArrangeCD (mode)

End Sub

Sub cmdExceltoPt() '�i�IDraw point data to CAD

Dim obj As New clsPlanMap

obj.ImportDataToCAD

End Sub

Sub cmdCreatePt() '���I

Dim obj As New clsPlanMap

obj.CreatePointByUser

End Sub

Sub cmdPtToLine() '�s�uplot line and specific feature

Dim obj As New clsPlanMap

obj.DrawLine (1)

MsgBox "�s�u����!", vbInformation

End Sub


Sub cmdCADPtToExcelPt() '��z�I��CAD point data to Excel point data

Dim obj As New clsPlanMap

obj.DealExportData

End Sub

Sub cmdSetDefaultFeature() 'set�w�]CD

Dim obj As New clsPlanMap

obj.SetDefaultFeature

End Sub

Sub cmdChangePt() '�I�����or����

Dim obj As New clsPTs_Table 'clsPlanMap

mode = InputBox("�п�ܲ��ʤ覡" & vbNewLine & "1.����" & vbNewLine & "2.����" & vbNewLine & "3.���")

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
    
    Set sht = Sheets("�`��")
    
    Call myfunc.ClearData(sht, 2, 1, 5)
    
    'FilePath = "G:\�ڪ����ݵw��\CADVBA\�����Ͻҵ{���\20190328.asc"
    
    If filePath = "" Then filePath = Application.GetOpenFilename
    
    If filePath = "False" Then MsgBox "������ɮ�!", vbCritical: End

    Open filePath For Input As #1
    Do Until EOF(1)
        Line Input #1, Line
        FileContent = FileContent & Line & vbCrLf
    Loop
    Close #1

    Lines = Split(FileContent, vbCrLf)
    
    prompt = Lines(0)
    
    mode = InputBox(prompt & vbNewLine & vbNewLine & "�п�ܼҦ�:" & vbNewLine & "1--P,E,N,Z,CD" & vbNewLine & "2--P,N,E,Z,CD", , "1")
    
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

mode = CAD.GetString("�п�ܼҦ�" & vbNewLine & "1.����" & vbNewLine & "2.��q")

If mode = 1 Then
    obj.interpolatePLCoorMix
ElseIf mode = 2 Then
    obj.interpolatePLCoor
End If

End Sub
