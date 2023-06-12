Attribute VB_Name = "useless"

'�쥻�p�e�Ψӥͦ�Skectup�W���I���>>�n�Ӱ��a��

Sub test_SKP()

With Sheets("�`��")

Dim vertices() As Double

lr = 10 ' .Cells(2, "G").End(xlDown).Row

cnt = lr - 2

ReDim vertices((cnt + 1) * 3 - 1)

For r = 2 To lr

    vertices(r_cnt) = .Cells(r, "H") ' - .Cells(2, "H")
    vertices(r_cnt + 1) = .Cells(r, "I") '- .Cells(2, "I")
    vertices(r_cnt + 2) = .Cells(r, "J")  '- .Cells(2, "J")

    r_cnt = r_cnt + 3
    
Next

End With

Call get3DCodeToSKP(vertices)

End Sub

Sub get3DCodeToSKP(ByVal vertices)

Dim coll As New Collection

'��l�Ƥ~�ݭn

txt = txt & "Model = Sketchup.active_model" & vbNewLine
txt = txt & "entities = Model.active_entities" & vbNewLine

'��l�Ƽƭ�

X0 = 0 '204238
Y0 = 0 ' 2622688
z0 = 0 ' 53

cross = 100 '�W�j�U�Q��

For i = 3 To UBound(vertices) Step 3

    cnt = cnt + 1

    X1 = vertices(i - 3) - X0
    Y1 = vertices(i + 1 - 3) - Y0
    Z1 = vertices(i + 2 - 3) - z0
    
    X2 = vertices(i) - X0
    Y2 = vertices(i + 1) - Y0
    Z2 = vertices(i + 2) - z0
    
    code1 = "point1 = Geom::Point3d.new(" & Int(X1) * cross & "," & Int(Y1) * cross & "," & Int(Z1) * cross & ")" & vbNewLine
    code2 = "point2 = Geom::Point3d.new(" & Int(X2) * cross & "," & Int(Y2) * cross & "," & Int(Z2) * cross & ")" & vbNewLine
    
    txt = txt & code1
    txt = txt & code2
    'txt = txt & "Line" & cnt & "= entities.add_line point1, point2" & vbNewLine
    txt = txt & "Line" & Int(Rnd() * 100000000) & cnt & "= entities.add_line point1, point2" & vbNewLine
    
    '�C��Line���W�r�����୫�ơA���M�|����
    
Next

Debug.Print txt
    
End Sub
