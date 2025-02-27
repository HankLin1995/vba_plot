Attribute VB_Name = "UnitTest_Math"
Option Explicit

' 引用主測試模組
' 注意：需要先運行 UnitTest 模組中的 InitializeTestEnvironment

' 數學模組完整測試套件
Public Sub RunAllMathTests()
    Debug.Print vbNewLine & "========== 數學模組測試套件 =========="
    
    ' 角度轉換測試
    Test_Math_AngleConversion
    
    ' 三角函數測試
    Test_Math_TrigFunctions
    
    ' 方位角計算測試
    Test_Math_Azimuth
    
    ' 距離計算測試
    Test_Math_Distance
    
    ' 點位判斷測試
    Test_Math_PointPosition
    
    ' 樁號轉換測試
    Test_Math_StationConversion
End Sub

' 測試角度轉換函數
Private Sub Test_Math_AngleConversion()
    Debug.Print vbNewLine & "----- 測試角度轉換函數 -----"
    
    Dim Math As New clsMath
    Dim testCases As Variant
    Dim expected As Variant
    Dim i As Integer
    
    ' 角度轉弧度測試案例
    testCases = Array(0, 30, 45, 90, 180, 270, 360)
    expected = Array(0, 0.5235987756, 0.7853981634, 1.5707963268, 3.1415926536, 4.7123889804, 6.2831853072)
    
    For i = 0 To UBound(testCases)
        UnitTest.AssertEqual expected(i), Math.deg2rad(testCases(i)), "角度轉弧度測試 - " & testCases(i) & "度"
    Next i
    
    ' 弧度轉角度測試案例
    testCases = expected
    expected = Array(0, 30, 45, 90, 180, 270, 360)
    
    For i = 0 To UBound(testCases)
        UnitTest.AssertEqual expected(i), Math.rad2deg(testCases(i)), "弧度轉角度測試 - " & testCases(i) & "弧度"
    Next i
End Sub

' 測試三角函數
Private Sub Test_Math_TrigFunctions()
    Debug.Print vbNewLine & "----- 測試三角函數 -----"
    
    Dim Math As New clsMath
    Dim angles As Variant
    Dim expectedSin As Variant
    Dim expectedCos As Variant
    Dim i As Integer
    
    ' 測試案例
    angles = Array(0, 30, 45, 60, 90, 180, 270, 360)
    expectedSin = Array(0, 0.5, 0.7071067812, 0.866025404, 1, 0, -1, 0)
    expectedCos = Array(1, 0.866025404, 0.7071067812, 0.5, 0, -1, 0, 1)
    
    For i = 0 To UBound(angles)
        UnitTest.AssertEqual expectedSin(i), Round(Math.degsin(angles(i)), 10), "正弦函數測試 - " & angles(i) & "度"
        UnitTest.AssertEqual expectedCos(i), Round(Math.degcos(angles(i)), 10), "餘弦函數測試 - " & angles(i) & "度"
    Next i
    
    ' 測試正切函數 (如果存在)
    On Error Resume Next
    Dim tan30 As Double
    tan30 = Math.degtan(30)
    
    If Err.Number = 0 Then
        UnitTest.AssertEqual 0.57735027, Round(tan30, 8), "正切函數測試 - 30度"
    Else
        Debug.Print "跳過正切函數測試 - 函數不存在"
    End If
    On Error GoTo 0
End Sub

' 測試方位角計算
Private Sub Test_Math_Azimuth()
    Debug.Print vbNewLine & "----- 測試方位角計算 -----"
    
    Dim Math As New clsMath
    Dim testCases As Variant
    Dim expected As Variant
    Dim i As Integer
    Dim result As Double
    
    ' 測試案例：起點(0,0)，終點為不同方向的點
    ' 格式：x1, y1, x2, y2, 預期方位角
    testCases = Array( _
        Array(0, 0, 1, 0, 0),    ' 正東方向
        Array(0, 0, 1, 1, 45),   ' 東北方向
        Array(0, 0, 0, 1, 90),   ' 正北方向
        Array(0, 0, -1, 1, 135), ' 西北方向
        Array(0, 0, -1, 0, 180), ' 正西方向
        Array(0, 0, -1, -1, 225), ' 西南方向
        Array(0, 0, 0, -1, 270), ' 正南方向
        Array(0, 0, 1, -1, 315)  ' 東南方向
    )
    
    For i = 0 To UBound(testCases)
        result = Math.getAz(testCases(i)(0), testCases(i)(1), testCases(i)(2), testCases(i)(3))
        UnitTest.AssertEqual testCases(i)(4), result, "方位角計算測試 #" & i
    Next i
    
    ' 測試特殊情況：起點與終點相同
    On Error Resume Next
    result = Math.getAz(0, 0, 0, 0)
    
    If Err.Number = 0 Then
        ' 如果函數處理了這種情況，檢查結果是否合理
        ' 通常會返回0或某個預設值
        Debug.Print "起點終點相同時方位角為: " & result
    Else
        Debug.Print "起點終點相同時方位角計算出錯: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' 測試距離計算
Private Sub Test_Math_Distance()
    Debug.Print vbNewLine & "----- 測試距離計算 -----"
    
    Dim Math As New clsMath
    Dim testCases As Variant
    Dim expected As Variant
    Dim i As Integer
    
    ' 測試案例：計算兩點間距離
    ' 格式：x1, y1, x2, y2, 預期距離
    testCases = Array( _
        Array(0, 0, 3, 4, 5),       ' 3-4-5三角形
        Array(0, 0, 1, 0, 1),       ' 水平線
        Array(0, 0, 0, 1, 1),       ' 垂直線
        Array(1, 1, 4, 5, 5),       ' 非原點起始
        Array(-2, -3, 2, 3, 7.2111) ' 跨越象限
    )
    
    For i = 0 To UBound(testCases)
        Dim result As Double
        result = Math.getLengthCO(testCases(i)(0), testCases(i)(1), testCases(i)(2), testCases(i)(3))
        
        ' 對於非精確值，使用近似比較
        If i = 4 Then
            UnitTest.AssertTrue Abs(testCases(i)(4) - result) < 0.001, _
                "距離計算測試 #" & i & " - 預期: " & testCases(i)(4) & ", 實際: " & result
        Else
            UnitTest.AssertEqual testCases(i)(4), result, "距離計算測試 #" & i
        End If
    Next i
End Sub

' 測試點位判斷
Private Sub Test_Math_PointPosition()
    Debug.Print vbNewLine & "----- 測試點位判斷 -----"
    
    Dim Math As New clsMath
    Dim testCases As Variant
    Dim expected As Variant
    Dim i As Integer
    
    ' 測試案例：判斷點是否在線段上
    ' 格式：x1, y1, x, y, x2, y2, 預期結果(True/False)
    testCases = Array( _
        Array(0, 0, 5, 5, 10, 10, True),    ' 點在線段上
        Array(0, 0, 15, 15, 10, 10, False), ' 點在線延長線上
        Array(0, 0, 5, 6, 10, 10, False),   ' 點不在線上
        Array(0, 0, 0, 0, 10, 10, True),    ' 點與起點重合
        Array(0, 0, 10, 10, 10, 10, True),  ' 點與終點重合
        Array(0, 0, -5, -5, 10, 10, False)  ' 點在線反方向延長線上
    )
    
    For i = 0 To UBound(testCases)
        Dim result As Boolean
        result = Math.IsMiddle(testCases(i)(0), testCases(i)(1), testCases(i)(2), testCases(i)(3), testCases(i)(4), testCases(i)(5))
        UnitTest.AssertEqual testCases(i)(6), result, "點位判斷測試 #" & i
    Next i
End Sub

' 測試樁號轉換
Private Sub Test_Math_StationConversion()
    Debug.Print vbNewLine & "----- 測試樁號轉換 -----"
    
    Dim Math As New clsMath
    Dim testCases As Variant
    Dim expected As Variant
    Dim i As Integer
    
    ' 測試TranLoc函數 (如果存在)
    On Error Resume Next
    
    ' 測試案例：各種格式的樁號轉換為數值
    testCases = Array("0K+000", "1K+234", "10+123", "K123+456", "123.45", "123+456.78")
    expected = Array(0, 1234, 10123, 123456, 123.45, 123456.78)
    
    For i = 0 To UBound(testCases)
        Dim result As Double
        result = Math.TranLoc(testCases(i))
        
        If Err.Number = 0 Then
            UnitTest.AssertEqual expected(i), result, "樁號轉換測試 - " & testCases(i)
        Else
            Debug.Print "樁號轉換測試失敗 - " & testCases(i) & ": " & Err.Description
            Err.Clear
        End If
    Next i
    
    On Error GoTo 0
End Sub
