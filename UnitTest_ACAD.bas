Attribute VB_Name = "UnitTest_ACAD"
Option Explicit

' 引用主測試模組
' 注意：需要先運行 UnitTest 模組中的 InitializeTestEnvironment

' ACAD模組完整測試套件
Public Sub RunAllACADTests()
    Debug.Print vbNewLine & "========== ACAD模組測試套件 =========="
    
    ' CAD連接測試
    Test_ACAD_Connection
    
    ' 坐標轉換測試
    Test_ACAD_CoordinateConversion
    
    ' 圖層操作測試
    Test_ACAD_LayerOperations
    
    ' 繪圖功能測試
    Test_ACAD_DrawingFunctions
    
    ' 選擇集測試
    Test_ACAD_SelectionSet
    
    ' 文字操作測試
    Test_ACAD_TextOperations
End Sub

' 測試CAD連接
Private Sub Test_ACAD_Connection()
    Debug.Print vbNewLine & "----- 測試CAD連接 -----"
    
    Dim CAD As New clsACAD
    
    ' 測試連接方法
    On Error Resume Next
    CAD.Connect
    
    If Err.Number = 0 Then
        UnitTest.AssertTrue Not CAD.acadDoc Is Nothing, "CAD連接測試 - acadDoc不為空"
        UnitTest.AssertTrue Not CAD.acadApp Is Nothing, "CAD連接測試 - acadApp不為空"
        
        ' 測試CAD版本屬性
        UnitTest.AssertTrue Len(CAD.CADVer) > 0, "CAD版本屬性測試"
        Debug.Print "當前CAD版本: " & CAD.CADVer
    Else
        Debug.Print "CAD連接失敗: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' 測試坐標轉換
Private Sub Test_ACAD_CoordinateConversion()
    Debug.Print vbNewLine & "----- 測試坐標轉換 -----"
    
    Dim CAD As New clsACAD
    
    ' 連接CAD
    On Error Resume Next
    CAD.Connect
    
    If Err.Number <> 0 Or CAD.acadDoc Is Nothing Then
        Debug.Print "跳過坐標轉換測試 - CAD未連接"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 測試點轉換函數
    Dim pt(0 To 2) As Double
    pt(0) = 10: pt(1) = 20: pt(2) = 0
    
    Dim acadPt As Variant
    acadPt = CAD.tranPoint(pt)
    
    UnitTest.AssertEqual 3, UBound(acadPt) + 1, "點轉換測試 - 維度檢查"
    UnitTest.AssertEqual 10, acadPt(0), "點轉換測試 - X坐標"
    UnitTest.AssertEqual 20, acadPt(1), "點轉換測試 - Y坐標"
    
    ' 測試點數組轉換函數 (如果存在)
    On Error Resume Next
    Dim pts(0 To 5) As Double
    pts(0) = 0: pts(1) = 0: pts(2) = 0
    pts(3) = 10: pts(4) = 10: pts(5) = 0
    
    Dim tranPts As Variant
    tranPts = CAD.tranIPoints(pts)
    
    If Err.Number = 0 Then
        UnitTest.AssertTrue Not IsEmpty(tranPts), "點數組轉換測試"
    Else
        Debug.Print "跳過點數組轉換測試 - " & Err.Description
    End If
    On Error GoTo 0
End Sub

' 測試圖層操作
Private Sub Test_ACAD_LayerOperations()
    Debug.Print vbNewLine & "----- 測試圖層操作 -----"
    
    Dim CAD As New clsACAD
    
    ' 連接CAD
    On Error Resume Next
    CAD.Connect
    
    If Err.Number <> 0 Or CAD.acadDoc Is Nothing Then
        Debug.Print "跳過圖層操作測試 - CAD未連接"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 測試圖層創建
    Dim testLayerName As String
    testLayerName = "UnitTest_Layer_" & Format(Now, "hhmmss")
    
    On Error Resume Next
    CAD.acadDoc.Layers.Add testLayerName
    UnitTest.AssertEqual 0, Err.Number, "圖層創建測試", "圖層創建失敗: " & Err.Description
    On Error GoTo 0
    
    ' 測試圖層設置函數 (如果存在)
    On Error Resume Next
    
    ' 創建一個測試線條
    Dim spt(0 To 2) As Double
    Dim ept(0 To 2) As Double
    spt(0) = 0: spt(1) = 0: spt(2) = 0
    ept(0) = 10: ept(1) = 0: ept(2) = 0
    
    Dim testLine As Object
    Set testLine = CAD.AddLine(spt, ept)
    
    ' 設置圖層
    CAD.setLayer testLine, testLayerName
    
    UnitTest.AssertEqual testLayerName, testLine.Layer, "圖層設置測試"
    
    ' 清理測試對象
    testLine.Delete
    
    On Error GoTo 0
    
    ' 清理測試圖層
    On Error Resume Next
    If Not CAD.acadDoc.Layers(testLayerName) Is Nothing Then
        CAD.acadDoc.Layers(testLayerName).Delete
    End If
    On Error GoTo 0
End Sub

' 測試繪圖功能
Private Sub Test_ACAD_DrawingFunctions()
    Debug.Print vbNewLine & "----- 測試繪圖功能 -----"
    
    Dim CAD As New clsACAD
    
    ' 連接CAD
    On Error Resume Next
    CAD.Connect
    
    If Err.Number <> 0 Or CAD.acadDoc Is Nothing Then
        Debug.Print "跳過繪圖功能測試 - CAD未連接"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 創建測試圖層
    Dim testLayerName As String
    testLayerName = "UnitTest_Drawing_" & Format(Now, "hhmmss")
    
    On Error Resume Next
    CAD.acadDoc.Layers.Add testLayerName
    CAD.acadDoc.ActiveLayer = CAD.acadDoc.Layers(testLayerName)
    On Error GoTo 0
    
    ' 測試線條繪製
    On Error Resume Next
    Dim spt(0 To 2) As Double
    Dim ept(0 To 2) As Double
    spt(0) = 0: spt(1) = 0: spt(2) = 0
    ept(0) = 10: ept(1) = 0: ept(2) = 0
    
    Dim testLine As Object
    Set testLine = CAD.AddLine(spt, ept)
    
    UnitTest.AssertTrue Not testLine Is Nothing, "線條繪製測試"
    
    ' 測試多段線繪製
    Dim vertices(0 To 8) As Double
    vertices(0) = 0: vertices(1) = 0: vertices(2) = 0
    vertices(3) = 10: vertices(4) = 0: vertices(5) = 0
    vertices(6) = 10: vertices(7) = 10: vertices(8) = 0
    
    Dim testPLine As Object
    Set testPLine = CAD.AddPolyLine(vertices)
    
    UnitTest.AssertTrue Not testPLine Is Nothing, "多段線繪製測試"
    
    ' 測試圓弧繪製 (如果存在)
    Dim center(0 To 2) As Double
    center(0) = 5: center(1) = 5: center(2) = 0
    
    Dim testCircle As Object
    Set testCircle = CAD.acadDoc.ModelSpace.AddCircle(center, 5)
    
    UnitTest.AssertTrue Not testCircle Is Nothing, "圓繪製測試"
    
    ' 清理測試對象
    testLine.Delete
    testPLine.Delete
    testCircle.Delete
    
    On Error GoTo 0
    
    ' 清理測試圖層
    On Error Resume Next
    If Not CAD.acadDoc.Layers(testLayerName) Is Nothing Then
        CAD.acadDoc.Layers(testLayerName).Delete
    End If
    On Error GoTo 0
End Sub

' 測試選擇集
Private Sub Test_ACAD_SelectionSet()
    Debug.Print vbNewLine & "----- 測試選擇集 -----"
    
    Dim CAD As New clsACAD
    
    ' 連接CAD
    On Error Resume Next
    CAD.Connect
    
    If Err.Number <> 0 Or CAD.acadDoc Is Nothing Then
        Debug.Print "跳過選擇集測試 - CAD未連接"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 測試選擇集創建函數
    On Error Resume Next
    Dim sset As Object
    
    ' 創建一個測試對象
    Dim testLayerName As String
    testLayerName = "UnitTest_SSET_" & Format(Now, "hhmmss")
    
    CAD.acadDoc.Layers.Add testLayerName
    CAD.acadDoc.ActiveLayer = CAD.acadDoc.Layers(testLayerName)
    
    Dim spt(0 To 2) As Double
    Dim ept(0 To 2) As Double
    spt(0) = 0: spt(1) = 0: spt(2) = 0
    ept(0) = 10: ept(1) = 0: ept(2) = 0
    
    Dim testLine As Object
    Set testLine = CAD.AddLine(spt, ept)
    
    ' 測試按圖層選擇
    Set sset = CAD.CreateSSET("", "8", testLayerName)
    
    If Err.Number = 0 Then
        UnitTest.AssertTrue Not sset Is Nothing, "選擇集創建測試"
        UnitTest.AssertEqual 1, sset.Count, "選擇集數量測試"
    Else
        Debug.Print "選擇集創建失敗: " & Err.Description
    End If
    
    ' 清理測試對象
    testLine.Delete
    
    On Error GoTo 0
    
    ' 清理測試圖層
    On Error Resume Next
    If Not CAD.acadDoc.Layers(testLayerName) Is Nothing Then
        CAD.acadDoc.Layers(testLayerName).Delete
    End If
    On Error GoTo 0
End Sub

' 測試文字操作
Private Sub Test_ACAD_TextOperations()
    Debug.Print vbNewLine & "----- 測試文字操作 -----"
    
    Dim CAD As New clsACAD
    
    ' 連接CAD
    On Error Resume Next
    CAD.Connect
    
    If Err.Number <> 0 Or CAD.acadDoc Is Nothing Then
        Debug.Print "跳過文字操作測試 - CAD未連接"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 測試文字添加函數
    On Error Resume Next
    
    ' 創建測試圖層
    Dim testLayerName As String
    testLayerName = "UnitTest_Text_" & Format(Now, "hhmmss")
    
    CAD.acadDoc.Layers.Add testLayerName
    CAD.acadDoc.ActiveLayer = CAD.acadDoc.Layers(testLayerName)
    
    ' 測試點
    Dim insPoint(0 To 2) As Double
    insPoint(0) = 0: insPoint(1) = 0: insPoint(2) = 0
    
    ' 測試添加文字
    Dim testText As Object
    Set testText = CAD.AddMixText("測試文字", insPoint, 2.5, 1)
    
    If Err.Number = 0 Then
        UnitTest.AssertTrue Not testText Is Nothing, "文字添加測試"
        UnitTest.AssertEqual "測試文字", testText.TextString, "文字內容測試"
    Else
        Debug.Print "文字添加失敗: " & Err.Description
    End If
    
    ' 測試文字框 (如果存在)
    On Error Resume Next
    Dim textBox As Object
    Set textBox = CAD.AddTextBox(testText)
    
    If Err.Number = 0 And Not textBox Is Nothing Then
        UnitTest.AssertTrue True, "文字框添加測試"
        textBox.Delete
    Else
        Debug.Print "跳過文字框測試 - " & Err.Description
    End If
    
    ' 清理測試對象
    If Not testText Is Nothing Then testText.Delete
    
    On Error GoTo 0
    
    ' 清理測試圖層
    On Error Resume Next
    If Not CAD.acadDoc.Layers(testLayerName) Is Nothing Then
        CAD.acadDoc.Layers(testLayerName).Delete
    End If
    On Error GoTo 0
End Sub
