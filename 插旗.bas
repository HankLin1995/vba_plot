Attribute VB_Name = "插旗"
Sub test_getBoundingBox()

    '適用於第一象限

    Dim CAD As New clsACAD
    Dim L As Double
    Dim H As Double

    Set sset2 = CAD.CreateSSET()
    Set entobj = sset2(0)
    
    rotate_degree = entobj.Rotation * 180 / 3.14
    
    Call entobj.GetBoundingBox(Min, Max)
    
    X0 = Min(0)
    Y0 = Min(1)
    X1 = Max(0)
    Y1 = Max(1)
    
    Call CAD.AddPoint(Min)
    Call CAD.AddPoint(Max)
    
    '---一元二次聯立方程式---
    
    a1 = Sin(rotate_degree / 180 * 3.14)
    b1 = Cos(rotate_degree / 180 * 3.14)
    c1 = Y1 - Y0
    
    a2 = Cos(rotate_degree / 180 * 3.14)
    b2 = Sin(rotate_degree / 180 * 3.14)
    c2 = X1 - X0
    
    Call SolveLinearEquations(a1, b1, c1, a2, b2, c2, L, H)
    
    '------底線------
    
    Xpt_LD = X0 + H * Sin(rotate_degree / 180 * 3.14)
    Ypt_LD = Y0
    
    Xpt_RD = X1
    Ypt_RD = Y1 - H * Cos(rotate_degree / 180 * 3.14)
    
    Call CAD.AddLineCO(Xpt_LD, Ypt_LD, Xpt_RD, Ypt_RD)
    
    '-------頂線------
    
    Xpt_LU = X0
    Ypt_LU = Y0 + H * Cos(rotate_degree / 180 * 3.14)
    
    Xpt_RU = X1 - H * Sin(rotate_degree / 180 * 3.14)
    Ypt_RU = Y1
    
    Call CAD.AddLineCO(Xpt_LU, Ypt_LU, Xpt_RU, Ypt_RU)

    
End Sub

Function SolveLinearEquations(ByVal a1 As Double, ByVal b1 As Double, ByVal c1 As Double, ByVal a2 As Double, ByVal b2 As Double, ByVal c2 As Double, ByRef x As Double, ByRef y As Double) As Boolean
    Dim determinant As Double
    
    determinant = a1 * b2 - a2 * b1
    
    If determinant = 0 Then
        ' 方程式?解
        SolveLinearEquations = False
    Else
        x = (c1 * b2 - c2 * b1) / determinant
        y = (a1 * c2 - a2 * c1) / determinant
        SolveLinearEquations = True
        
    End If
    
End Function
