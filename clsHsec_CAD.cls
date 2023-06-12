VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHsec_CAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===2021/11/26 NOTE===
'===Author:HankLin===

'getEGs >> �q���_���Ψ��o���˽u
'getEGs_ICAD >> �q���_���Ψ��o���˽u(FOR ICAD)
'getLocFromEG
'getPLs
'getEGTable
'AppenData
'SortData
'openLayers

Private sht_EG As Object
Private sht_CL As Object
Private CAD As New clsACAD
Private myFunc As New clsMyfunction

Private collEGs As New Collection
Private collPLs As New Collection

Const CL_CROSSLINE_LAYER_NAME = "������-���_��L"

Private Sub Class_Initialize()

Set sht_EG = Sheets("���_��")
Set sht_CL = Sheets("���߽u")

End Sub

Sub getEGs()

If CAD.CADVer = "ICAD" Then Call getEGs_ICAD: Exit Sub

Set EGs = CAD.CreateSSET("CrossLine", "0,8", "line," & CL_CROSSLINE_LAYER_NAME) '���_�����˽u

For Each EG In EGs

    collEGs.Add EG

Next

Set collEGs = myFunc.ReverseColl(collEGs) '�ͦ��ɶ����ߥ��Q����


End Sub

Private Sub getEGs_ICAD()

'ProgeCAD�ݭn�S�O�B�z�ϼh���D

Call openLayers(CL_CROSSLINE_LAYER_NAME)

Set EGs = CAD.CreateSSET("CrossLine", "0,8", "line," & CL_CROSSLINE_LAYER_NAME) '���_�����˽u

Call openLayers

Dim ar(): ReDim ar(EGs.Count - 1, 1)

For Each EG In EGs

    ar(j, 0) = getLocFromEG(EG)
    ar(j, 1) = EG
    
    j = j + 1

Next

myFunc.MergeSort_array ar

For i = LBound(ar) To UBound(ar)

    collEGs.Add CAD.acaddoc.HandleToObject(ar(i, 1))

Next

End Sub

Private Function getLocFromEG(ByVal EG)

spt = CAD.tranIPoint(EG.startpoint)

With sht_CL 'Sheets("���߽u")

Set rng = .Columns(2).Find(spt(0))

r = rng.Row

getLocFromEG = .Cells(r, 1)

End With

End Function

Sub getPLs()

CAD.acaddoc.Layers("������-���_��").LayerOn = False
CAD.acaddoc.Layers("������-���_��L").LayerOn = False
CAD.acaddoc.Layers("������-�I").LayerOn = False
CAD.acaddoc.Layers("������-�I�W").LayerOn = False
CAD.acaddoc.Layers("������-�I���{").LayerOn = False
CAD.acaddoc.Layers("������-�I�s�X").LayerOn = False

Set PLs = CAD.CreateSSET("SS2")

For Each PL In PLs

    If TypeName(PL) <> "IAcadBlockReference" Then
        collPLs.Add PL
    End If

Next

CAD.acaddoc.Layers("������-���_��").LayerOn = True
CAD.acaddoc.Layers("������-���_��L").LayerOn = True
CAD.acaddoc.Layers("������-�I").LayerOn = True
CAD.acaddoc.Layers("������-�I�W").LayerOn = True
CAD.acaddoc.Layers("������-�I���{").LayerOn = True
CAD.acaddoc.Layers("������-�I�s�X").LayerOn = True

End Sub

Sub getEGTable(ByVal IsFromTable As Boolean)

Dim IsIntersect As Boolean
Dim IsFirst As Boolean
Dim loc As Variant
Dim pts As New clsPTs_Table
Set collPT_Table = pts.getCollPointTable

Call myFunc.ClearData(sht_EG, 2, 1, 4)

For Each EG In collEGs
    
    arr = CAD.tranIPoint(EG.startpoint)
    X0 = arr(0): Y0 = arr(1)

    IsFirst = True '�_�l�I�P�_

    For Each PL In collPLs
        
        Dim PLobj As New clsPL
        Call PLobj.getPropertiesByPL(PL)
        'crossPT = CAD.tranIPoints(EG.IntersectWith(PL, 1))  'acExtendThisEntity))
        crossPT = CAD.IntersectWith(EG, PL, 1, IsIntersect) '20210901 ���~����
        CD = PLobj.PLname
        
        If IsIntersect Then '20210901
        
            For i = 0 To UBound(crossPT) Step 3
    
                X1 = crossPT(i)
                Y1 = crossPT(i + 1)
    
                Distance = Sqr((X1 - X0) ^ 2 + (Y1 - Y0) ^ 2)
                Z = PLobj.interpolation(crossPT, IsFromTable, collPT_Table)
                
                Call AppenData(getLocFromEG(EG), Distance, Z, CD, IsFirst)
                
            Next

        End If

    Next
    
    Call SortData
    
Next

End Sub

Sub AppenData(ByVal loc As String, ByVal dx As Double, ByVal Z As Double, ByVal CD As String, ByRef IsFirst As Boolean) '�o�����٦����D

With sht_EG

If IsFirst = True Then

    IsFirst = False
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row + 1 + 1
    If lr = 3 Then lr = 2
    
    .Cells(lr, 4) = loc
    
Else
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
End If

.Cells(lr, 1) = dx
.Cells(lr, 2) = Z
.Cells(lr, 3) = CD

End With

End Sub

Sub SortData()

With sht_EG

    sr = .Cells(.Rows.Count, 1).End(xlUp).End(xlUp).Row
    
    If sr = 1 Then sr = 2
    
    er = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    .Range("A" & sr & ":C" & er).Sort key1:=.Range("A" & sr & ":A" & er), order1:=xlAscending
    

End With

End Sub

Sub openLayers(Optional openlayer As String = "")

If openlayer = "" Then

    For Each ly In CAD.acaddoc.Layers
        ly.LayerOn = True
    Next

Else

    open_layers = Split(openlayer, ",")

        For Each ly In CAD.acaddoc.Layers
        
            ly.LayerOn = False
            
            For Each open_layer In open_layers
            
                If ly.Name = open_layer Then ly.LayerOn = True
            
            Next
            
        Next

End If

End Sub



