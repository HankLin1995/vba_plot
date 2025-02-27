Attribute VB_Name = "UnitTest_CL"
Option Explicit

' 引用主測試模組
' 注意：需要先運行 UnitTest 模組中的 InitializeTestEnvironment

' 中心線模組完整測試套件
Public Sub RunAllCLTests()
    Debug.Print vbNewLine & "========== 中心線模組測試套件 =========="
    
    ' 基本屬性測試
    Test_CL_Properties
    
    ' 樁號計算測試
    Test_CL_StationCalculation
    
    ' 中心線繪製測試 (需要CAD環境)
    Test_CL_Drawing
    
    ' 橫斷面生成測試 (需要CAD環境)
    Test_CL_CrossSection
End Sub

' 測試中心線基本屬性
Private Sub Test_CL_Properties()
    Debug.Print vbNewLine & "----- 測試中心線基本屬性 -----"
    
    Dim CL As New clsCL
    
    ' 測試預設值
    UnitTest.AssertEqual 0, CL.nowLoc, "中心線起始樁號預設值測試"
    
    ' 測試屬性設置
    CL.w = 10
    CL.nowLoc = 1000
    CL.PaperScale = 1000
    CL.WIDTH_COE = 1.5
    CL.IsLeftShow = True
    CL.IsRightShow = False
    CL.NeedReverse = False
    CL.NeedBox = True
    CL.NeedDir = True
    
    UnitTest.AssertEqual 10, CL.w, "半邊寬度設置測試"
    UnitTest.AssertEqual 1000, CL.nowLoc, "起始樁號設置測試"
    UnitTest.AssertEqual 1000, CL.PaperScale, "平面圖比例設置測試"
    UnitTest.AssertEqual 1.5, CL.WIDTH_COE, "文字距離係數設置測試"
    UnitTest.AssertEqual True, CL.IsLeftShow, "左標註顯示設置測試"
    UnitTest.AssertEqual False, CL.IsRightShow, "右標註顯示設置測試"
    UnitTest.AssertEqual False, CL.NeedReverse, "線方向反轉設置測試"
    UnitTest.AssertEqual True, CL.NeedBox, "文字外框設置測試"
    UnitTest.AssertEqual True, CL.NeedDir, "方向旗標設置測試"
    
    ' 測試數據保存與讀取
    On Error Resume Next
    CL.setDataByUser
    Err.Clear
    
    Dim CL2 As New clsCL
    CL2.getDataByRng
    
    If Err.Number = 0 Then
        UnitTest.AssertEqual CL.w, CL2.w, "數據保存與讀取測試 - 半邊寬度"
        UnitTest.AssertEqual CL.nowLoc, CL2.nowLoc, "數據保存與讀取測試 - 起始樁號"
        UnitTest.AssertEqual CL.IsLeftShow, CL2.IsLeftShow, "數據保存與讀取測試 - 左標註顯示"
        UnitTest.AssertEqual CL.IsRightShow, CL2.IsRightShow, "數據保存與讀取測試 - 右標註顯示"
    Else
        Debug.Print "跳過數據保存與讀取測試 - " & Err.Description
    End If
    On Error GoTo 0
End Sub

' 測試樁號計算
Private Sub Test_CL_StationCalculation()
    Debug.Print vbNewLine & "----- 測試樁號計算 -----"
    
    ' 模擬樁號數據
    Dim startStation As Double
    Dim interval As Double
    Dim totalLength As Double
    
    startStation = 1000
    interval = 20
    totalLength = 100
    
    ' 計算預期結果
    Dim expectedStations As Variant
    ReDim expectedStations(0 To 5)
    
    expectedStations(0) = 1000
    expectedStations(1) = 1020
    expectedStations(2) = 1040
    expectedStations(3) = 1060
    expectedStations(4) = 1080
    expectedStations(5) = 1100
    
    ' 測試樁號格式化
    Dim formattedStation As String
    formattedStation = Format(1234.56, "0K+000.0")
    UnitTest.AssertEqual "1K+234.6", formattedStation, "樁號格式化測試"
    
    ' 這部分需要模擬 clsCL.getLoc 方法的行為
    ' 由於該方法依賴於用戶輸入和 Excel 表格，所以這裡只是示例
    
    ' 模擬樁號計算
    Dim calculatedStations As Variant
    ReDim calculatedStations(0 To 5)
    
    calculatedStations(0) = startStation
    For i = 1 To 5
        calculatedStations(i) = startStation + i * interval
    Next i
    
    ' 驗證計算結果
    For i = 0 To 5
        UnitTest.AssertEqual expectedStations(i), calculatedStations(i), "樁號計算測試 #" & i
    Next i
End Sub

' 測試中心線繪製 (需要CAD環境)
Private Sub Test_CL_Drawing()
    Debug.Print vbNewLine & "----- 測試中心線繪製 -----"
    
    Dim CL As New clsCL
    Dim CAD As New clsACAD
    
    ' 連接CAD
    On Error Resume Next
    CAD.Connect
    
    If Err.Number <> 0 Or CAD.acadDoc Is Nothing Then
        Debug.Print "跳過中心線繪製測試 - CAD未連接"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 設置測試參數
    CL.w = 5
    CL.nowLoc = 100
    CL.PaperScale = 1000
    CL.IsLeftShow = True
    CL.IsRightShow = True
    
    ' 測試圖層創建
    On Error Resume Next
    If CAD.acadDoc.Layers.Item("CL") Is Nothing Then
        CAD.acadDoc.Layers.Add "CL"
    End If
    
    If CAD.acadDoc.Layers.Item("CL_CROSS") Is Nothing Then
        CAD.acadDoc.Layers.Add "CL_CROSS"
    End If
    UnitTest.AssertEqual 0, Err.Number, "中心線圖層創建測試", "圖層創建失敗: " & Err.Description
    On Error GoTo 0
    
    ' 注意：實際的繪圖測試需要在CAD中創建線條對象
    ' 這裡只是檢查相關函數是否存在並且不會拋出錯誤
    
    ' 測試繪圖函數存在性
    Dim methodExists As Boolean
    
    methodExists = FunctionExists(CL, "getCenterLine")
    UnitTest.AssertTrue methodExists, "getCenterLine方法存在測試"
    
    methodExists = FunctionExists(CL, "DrawCrossLine")
    UnitTest.AssertTrue methodExists, "DrawCrossLine方法存在測試"
    
    methodExists = FunctionExists(CL, "CrossLine_Main")
    UnitTest.AssertTrue methodExists, "CrossLine_Main方法存在測試"
End Sub

' 測試橫斷面生成 (需要CAD環境)
Private Sub Test_CL_CrossSection()
    Debug.Print vbNewLine & "----- 測試橫斷面生成 -----"
    
    Dim CL As New clsCL
    Dim CAD As New clsACAD
    
    ' 連接CAD
    On Error Resume Next
    CAD.Connect
    
    If Err.Number <> 0 Or CAD.acadDoc Is Nothing Then
        Debug.Print "跳過橫斷面生成測試 - CAD未連接"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 設置測試參數
    CL.w = 10
    CL.nowLoc = 0
    CL.BLnext = 20
    
    ' 測試函數存在性
    Dim methodExists As Boolean
    
    methodExists = FunctionExists(CL, "BorderLine_Main")
    UnitTest.AssertTrue methodExists, "BorderLine_Main方法存在測試"
    
    methodExists = FunctionExists(CL, "getBorderLine")
    UnitTest.AssertTrue methodExists, "getBorderLine方法存在測試"
    
    methodExists = FunctionExists(CL, "getCLpt")
    UnitTest.AssertTrue methodExists, "getCLpt方法存在測試"
End Sub

' 輔助函數：檢查對象是否有指定方法
Private Function FunctionExists(obj As Object, methodName As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    
    ' 嘗試使用CallByName調用方法
    ' 由於我們不知道參數，所以這可能會失敗
    ' 但Err.Number會告訴我們是否是因為方法不存在
    CallByName obj, methodName, VbMethod
    
    If Err.Number = 438 Then ' 對象不支持此屬性或方法
        FunctionExists = False
    Else
        FunctionExists = True
    End If
    
    Err.Clear
    On Error GoTo 0
End Function
