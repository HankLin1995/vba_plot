Attribute VB_Name = "UnitTest"
Option Explicit

' 全局變量用於跟踪測試結果
Private m_totalTests As Integer
Private m_passedTests As Integer
Private m_failedTests As Integer
Private m_testResults As String

' 初始化測試環境
Public Sub InitializeTestEnvironment()
    m_totalTests = 0
    m_passedTests = 0
    m_failedTests = 0
    m_testResults = ""
    Debug.Print "開始單元測試..." & vbNewLine
End Sub

' 顯示測試結果
Public Sub ShowTestResults()
    Debug.Print vbNewLine & "測試結果摘要:"
    Debug.Print "總測試數: " & m_totalTests
    Debug.Print "通過測試: " & m_passedTests
    Debug.Print "失敗測試: " & m_failedTests
    Debug.Print "通過率: " & Format(m_passedTests / IIf(m_totalTests = 0, 1, m_totalTests), "0.0%")
    
    If m_failedTests > 0 Then
        Debug.Print vbNewLine & "失敗測試詳情:"
        Debug.Print m_testResults
    End If
End Sub

' 斷言函數 - 檢查條件是否為真
Public Sub AssertTrue(condition As Boolean, testName As String, Optional message As String = "")
    m_totalTests = m_totalTests + 1
    
    If condition Then
        m_passedTests = m_passedTests + 1
        Debug.Print "✓ 通過: " & testName
    Else
        m_failedTests = m_failedTests + 1
        Debug.Print "✗ 失敗: " & testName & IIf(Len(message) > 0, " - " & message, "")
        m_testResults = m_testResults & "✗ " & testName & IIf(Len(message) > 0, " - " & message, "") & vbNewLine
    End If
End Sub

' 斷言函數 - 檢查條件是否為假
Public Sub AssertFalse(condition As Boolean, testName As String, Optional message As String = "")
    AssertTrue Not condition, testName, message
End Sub

' 斷言函數 - 檢查兩個值是否相等
Public Sub AssertEqual(expected, actual, testName As String, Optional message As String = "")
    Dim isEqual As Boolean
    
    If IsObject(expected) And IsObject(actual) Then
        ' 對象比較 (僅檢查引用是否相同)
        isEqual = (expected Is actual)
    ElseIf IsArray(expected) And IsArray(actual) Then
        ' 數組比較
        isEqual = ArraysEqual(expected, actual)
    Else
        ' 值比較
        isEqual = (expected = actual)
    End If
    
    If Not isEqual Then
        message = IIf(Len(message) > 0, message & " - ", "") & _
                 "預期: [" & CStr(expected) & "], 實際: [" & CStr(actual) & "]"
    End If
    
    AssertTrue isEqual, testName, message
End Sub

' 輔助函數 - 比較兩個數組是否相等
Private Function ArraysEqual(arr1, arr2) As Boolean
    Dim i As Long
    
    ' 檢查數組維度
    If LBound(arr1) <> LBound(arr2) Or UBound(arr1) <> UBound(arr2) Then
        ArraysEqual = False
        Exit Function
    End If
    
    ' 比較每個元素
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) <> arr2(i) Then
            ArraysEqual = False
            Exit Function
        End If
    Next i
    
    ArraysEqual = True
End Function

' 運行所有測試
Public Sub RunAllTests()
    Call InitializeTestEnvironment
    
    ' 運行數學模組測試
    Call Test_Math_Module
    
    ' 運行ACAD模組測試
    Call Test_ACAD_Module
    
    ' 運行中心線模組測試
    Call Test_CL_Module
    
    ' 運行縱斷面模組測試
    Call Test_Longitudinal_Module
    
    ' 顯示測試結果
    Call ShowTestResults
End Sub

' 數學模組測試
Public Sub Test_Math_Module()
    Debug.Print vbNewLine & "===== 測試 clsMath 模組 ====="
    
    Dim Math As New clsMath
    
    ' 測試角度轉換
    AssertEqual 0.5235987756, Math.deg2rad(30), "角度轉弧度測試", "角度轉弧度計算不正確"
    AssertEqual 45, Math.rad2deg(0.7853981634), "弧度轉角度測試", "弧度轉角度計算不正確"
    
    ' 測試三角函數
    AssertEqual 0.5, Math.degsin(30), "正弦函數測試", "正弦函數計算不正確"
    AssertEqual 0.866025404, Math.degcos(30), "餘弦函數測試", "餘弦函數計算不正確"
    
    ' 測試方位角計算
    AssertEqual 45, Math.getAz(0, 0, 1, 1), "方位角測試 - 東北方向", "方位角計算不正確"
    AssertEqual 135, Math.getAz(0, 0, 1, -1), "方位角測試 - 東南方向", "方位角計算不正確"
    
    ' 測試距離計算
    AssertEqual 5, Math.getLengthCO(0, 0, 3, 4), "距離計算測試", "距離計算不正確"
    
    ' 測試點位判斷
    AssertTrue Math.IsMiddle(0, 0, 5, 5, 10, 10), "點位判斷測試 - 點在線上", "點位判斷不正確"
    AssertFalse Math.IsMiddle(0, 0, 15, 15, 10, 10), "點位判斷測試 - 點不在線上", "點位判斷不正確"
End Sub

' ACAD模組測試
Public Sub Test_ACAD_Module()
    Debug.Print vbNewLine & "===== 測試 clsACAD 模組 ====="
    
    Dim CAD As New clsACAD
    
    ' 測試CAD連接 (僅檢查是否拋出錯誤)
    On Error Resume Next
    CAD.Connect
    AssertEqual 0, Err.Number, "CAD連接測試", "CAD連接失敗: " & Err.Description
    On Error GoTo 0
    
    ' 如果CAD已連接，則進行更多測試
    If Not CAD.acadDoc Is Nothing Then
        ' 測試點轉換
        Dim pt(0 To 2) As Double
        pt(0) = 1: pt(1) = 2: pt(2) = 0
        
        Dim acadPt As Variant
        acadPt = CAD.tranPoint(pt)
        
        AssertEqual 3, UBound(acadPt) + 1, "點轉換測試 - 維度檢查", "點轉換後維度不正確"
        AssertEqual 1, acadPt(0), "點轉換測試 - X坐標", "點轉換後X坐標不正確"
        AssertEqual 2, acadPt(1), "點轉換測試 - Y坐標", "點轉換後Y坐標不正確"
        
        ' 測試圖層創建 (僅檢查是否拋出錯誤)
        Dim testLayerName As String
        testLayerName = "UnitTest_Layer"
        
        On Error Resume Next
        CAD.acadDoc.Layers.Add testLayerName
        AssertEqual 0, Err.Number, "圖層創建測試", "圖層創建失敗: " & Err.Description
        On Error GoTo 0
        
        ' 清理測試圖層
        On Error Resume Next
        If Not CAD.acadDoc.Layers(testLayerName) Is Nothing Then
            CAD.acadDoc.Layers(testLayerName).Delete
        End If
        On Error GoTo 0
    End If
End Sub

' 中心線模組測試
Public Sub Test_CL_Module()
    Debug.Print vbNewLine & "===== 測試 clsCL 模組 ====="
    
    Dim CL As New clsCL
    
    ' 測試參數設置
    CL.w = 5
    CL.nowLoc = 100
    CL.IsLeftShow = True
    CL.IsRightShow = True
    
    AssertEqual 5, CL.w, "參數設置測試 - 寬度", "寬度參數設置不正確"
    AssertEqual 100, CL.nowLoc, "參數設置測試 - 起始樁號", "起始樁號參數設置不正確"
    AssertEqual True, CL.IsLeftShow, "參數設置測試 - 左標註", "左標註參數設置不正確"
    AssertEqual True, CL.IsRightShow, "參數設置測試 - 右標註", "右標註參數設置不正確"
    
    ' 更多測試需要在CAD環境中進行
End Sub

' 縱斷面模組測試
Public Sub Test_Longitudinal_Module()
    Debug.Print vbNewLine & "===== 測試 縱斷面 模組 ====="
    
    Dim Long As New clsLongitudinal
    
    ' 測試基本參數設置
    On Error Resume Next
    Long.setScale 1000, 100
    AssertEqual 0, Err.Number, "縱斷面比例設置測試", "縱斷面比例設置失敗: " & Err.Description
    On Error GoTo 0
    
    ' 更多測試需要在CAD環境中進行
End Sub

' 測試樣例 - 檢查是否有特定圖塊
Sub test_checkIfBlockIn()
    Debug.Print vbNewLine & "===== 測試圖塊檢查函數 ====="
    
    Dim CAD As New clsACAD
    Dim blockName As String
    blockName = "TEST_BLOCK"
    
    ' 此測試需要在CAD環境中進行
    On Error Resume Next
    CAD.Connect
    If Err.Number = 0 And Not CAD.acadDoc Is Nothing Then
        Dim result As Boolean
        result = False ' 這裡應該是您的圖塊檢查函數
        
        AssertTrue result, "圖塊存在測試", "圖塊 " & blockName & " 不存在"
    Else
        Debug.Print "跳過圖塊測試 - CAD未連接"
    End If
    On Error GoTo 0
End Sub
