Attribute VB_Name = "MyFunction"
Function Eval(ByVal s As String)

'���h��F���i�ֳt�B�z,�s������@

Dim cal As String

For i = 1 To Len(s)

    ch = mid(s, i, 1)
    
    If IsNumeric(ch) Then '�P�_�O�_���Ʀr
        cal = cal + ch
    ElseIf ch = "(" Or ch = "[" Or ch = "{" Then  '�A��
        cal = cal + "("
    ElseIf ch = ")" Or ch = "]" Or ch = "}" Then    '�A��
        cal = cal + ")"
    ElseIf ch = "+" Or ch = "-" Or ch = "*" Or ch = "/" Then '�B���
        cal = cal + ch
    ElseIf ch = "." Then '��L����
        cal = cal + ch
    End If

Next

Eval = Application.Evaluate(cal)

End Function

Function getMeanH(ByVal myloc As Double)

Dim myMath As New clsMath

With Sheets("�a�_��LEVEL")

    Lc = .Cells(1, 1).End(xlToRight).Column
    
    For c = 2 To Lc - 1
    
        bc = .Cells(1, c)
        bh = .Cells(2, c)
        ec = .Cells(1, c + 1)
        eH = .Cells(2, c + 1)
        
        getMeanH = myMath.interpolation(myloc, bc, bh, ec, eH)
        
        If getMeanH <> 0 Then Exit Function
        
    Next

End With

End Function

Function TranLoc(ByVal Data As String) As Double

'�θ����A�ন�i�p�⤧�θ�

tmp = Split(Data, "+")

If UBound(tmp) = 0 Then TranLoc = CDbl(Data): Exit Function

tloc = tmp(0) '�d���
dloc = tmp(1)

If dloc Like "*(*" Then

    tmp2 = Split(dloc, "(")

    tmp3 = Split(tmp2(0), ".")

    dloc = tmp3(0) + tmp3(1) / 10
    
End If

For i = 1 To Len(tloc)

    ch = mid(tloc, i, 1)
    If IsNumeric(ch) Then ref = ref & ch

Next

TranLoc = CDbl(ref) * 1000 + CDbl(dloc)
    
End Function





