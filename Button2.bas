Attribute VB_Name = "Button2"
Sub plotAreaBorder()

Dim CAD As New clsACAD

Dim pt0(2) As Double
Dim pt1(2) As Double
Dim mpt(2) As Double

With Sheets("中心線")

For r = 3 To 21

    myloc = .Cells(r, 1)
    tmp = Split(.Cells(r, 3), ",")

    pt0(0) = tmp(0)
    pt0(1) = tmp(1)
    pt1(0) = tmp(2)
    pt1(1) = tmp(3)
    
    Set recobj = CAD.PlotRec(pt0, pt1)

    recobj.Layer = "XLINE"
    
    mpt(0) = (pt0(0) + pt1(0)) / 2
    mpt(1) = (pt0(1) + pt1(1)) / 2

    Set txtobj = CAD.AddText(myloc, mpt, 100 * 3)

    txtobj.Layer = "XLINE"

Next

End With

End Sub

Sub renewVport()

Dim coll As New Collection

With Sheets("圖說")

For r = 3 To .Cells(3, 1).End(xlDown).Row

    s = .Cells(r, "J") & ":" & .Cells(r, "C")

    mykey = .Cells(r, "C")

    coll.Add s, mykey
    
Next

End With

Dim CAD As New clsACAD

Set sset = CAD.CreateSSET("SS1", "8", "視埠")

For Each it In sset

    If TypeName(it) = "IAcadText" Then
    
        tmp = Split(it.TextString, ":")
        
        myfind = tmp(1)
    
        Debug.Print "o=" & it.TextString
        Debug.Print "n=" & coll(myfind)
        
        it.TextString = coll(myfind)
    
    End If

Next

End Sub

Sub defineAreaBorder() '更新橫斷面定義框位置

'TODO
'1.收集框框
'2.指定框框的名稱

Dim CAD As New clsACAD
Dim collXSECs As New Collection

Set XSECs = CAD.CreateSSET("XSEC", "8", "XLINE")

For Each XSEC In XSECs

    If TypeName(XSEC) = "IAcadPolyline" Then
    
        Call CAD.GetBoundingBox(XSEC, MinX, MinY, MaxX, MaxY)
        
        For Each XSEC2 In XSECs
        
            If TypeName(XSEC2) = "IAcadText" Then
                
                myloc = XSEC2.TextString
                pt = XSEC2.InsertionPoint
                
                If IsInvolvedFromBorder(pt(0), pt(1), MinX, MinY, MaxX, MaxY) Then
                
                    collXSECs.Add myloc & ":" & MinX & "," & MinY & "," & MaxX & "," & MaxY
                
                End If
        
            End If
        
        Next
    
    End If

Next

With Sheets("中心線")

    r = 3

    For Each it In collXSECs
    
        tmp = Split(it, ":")
    
        .Cells(r, "K") = tmp(0)
        .Cells(r, "L") = tmp(1)
    
        r = r + 1
    
    Next
    
End With

End Sub

Function IsInvolvedFromBorder(ByVal midX, ByVal midY, ByVal Border_minX, ByVal Border_minY, ByVal Border_maxX, ByVal Border_maxY)

IsInvolvedFromBorder = False

If midX >= Border_minX And midX <= Border_maxX And midY >= Border_minY And midY <= Border_maxY Then
    
    IsInvolvedFromBorder = True

End If

End Function

