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

'getEGs >> 從橫斷面樁取得取樣線
'getEGs_ICAD >> 從橫斷面樁取得取樣線(FOR ICAD)
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

Const CL_CROSSLINE_LAYER_NAME = "平面圖-橫斷樁L"

Private Sub Class_Initialize()

Set sht_EG = Sheets("橫斷面")
Set sht_CL = Sheets("中心線")

End Sub

Sub getEGs()

If CAD.CADVer = "ICAD" Then Call getEGs_ICAD: Exit Sub

Set EGs = CAD.CreateSSET("CrossLine", "0,8", "line," & CL_CROSSLINE_LAYER_NAME) '橫斷面取樣線

For Each EG In EGs

    collEGs.Add EG

Next

Set collEGs = myFunc.ReverseColl(collEGs) '生成時間較晚先被收集


End Sub

Private Sub getEGs_ICAD()

'ProgeCAD需要特別處理圖層問題

Call openLayers(CL_CROSSLINE_LAYER_NAME)

Set EGs = CAD.CreateSSET("CrossLine", "0,8", "line," & CL_CROSSLINE_LAYER_NAME) '橫斷面取樣線

Call openLayers

Dim ar(): ReDim ar(EGs.Count - 1, 1)

For Each EG In EGs

    ar(j, 0) = getLocFromEG(EG)
    ar(j, 1) = EG
    
    j = j + 1

Next

myFunc.MergeSort_array ar

For i = LBound(ar) To UBound(ar)

    collEGs.Add CAD.acadDoc.HandleToObject(ar(i, 1))

Next

End Sub

Private Function getLocFromEG(ByVal EG)

spt = CAD.tranIPoint(EG.startpoint)

With sht_CL 'Sheets("中心線")

Set rng = .Columns(2).Find(spt(0))

r = rng.Row

getLocFromEG = .Cells(r, 1)

End With

End Function

Sub getPLs()

CAD.acadDoc.Layers("平面圖-橫斷樁").LayerOn = False
CAD.acadDoc.Layers("平面圖-橫斷樁L").LayerOn = False
CAD.acadDoc.Layers("平面圖-點").LayerOn = False
CAD.acadDoc.Layers("平面圖-點名").LayerOn = False
CAD.acadDoc.Layers("平面圖-點高程").LayerOn = False
CAD.acadDoc.Layers("平面圖-點編碼").LayerOn = False

Set PLs = CAD.CreateSSET("SS2")

For Each PL In PLs

    'Debug.Print TypeName(PL)
    If TypeName(PL) Like "*Poly*" Then
        collPLs.Add PL
    End If

Next

CAD.acadDoc.Layers("平面圖-橫斷樁").LayerOn = True
CAD.acadDoc.Layers("平面圖-橫斷樁L").LayerOn = True
CAD.acadDoc.Layers("平面圖-點").LayerOn = True
CAD.acadDoc.Layers("平面圖-點名").LayerOn = True
CAD.acadDoc.Layers("平面圖-點高程").LayerOn = True
CAD.acadDoc.Layers("平面圖-點編碼").LayerOn = True

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

    IsFirst = True '起始點判斷

    For Each PL In collPLs
        
        Dim PLobj As New clsPL
        Call PLobj.getPropertiesByPL(PL)
        'crossPT = CAD.tranIPoints(EG.IntersectWith(PL, 1))  'acExtendThisEntity))
        crossPT = CAD.IntersectWith(EG, PL, 1, IsIntersect) '20210901 錯誤偵測
        CD = PLobj.PLname
       ' kk = kk + 1
        'Debug.Print CD & kk
       ' If kk = 134 Then Stop
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

Sub AppenData(ByVal loc As String, ByVal dx As Double, ByVal Z As Double, ByVal CD As String, ByRef IsFirst As Boolean) '這部分還有問題

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

    For Each ly In CAD.acadDoc.Layers
        ly.LayerOn = True
    Next

Else

    open_layers = Split(openlayer, ",")

        For Each ly In CAD.acadDoc.Layers
        
            ly.LayerOn = False
            
            For Each open_layer In open_layers
            
                If ly.Name = open_layer Then ly.LayerOn = True
            
            Next
            
        Next

End If

End Sub




