VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMyfunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function getBlankColl(ByVal sht As Object, ByVal col As Integer)

'20200729
'總表的線分段資料
'橫斷面的分段資料

Dim collBlank As New Collection

collBlank.Add 1

With sht

    lr = .Cells(.Rows.Count, col).End(xlUp).Row

    For r = 2 To lr
    
        If .Cells(r, col) = "" Then collBlank.Add r
    
    Next
    
    collBlank.Add lr + 1
    
End With

Set getBlankColl = collBlank

End Function

Sub ClearData(ByVal sht As Object, ByVal sr As Integer, ByVal col As Integer, ByVal xrange As Integer)

With sht

    lr = .Cells(.Rows.Count, col).End(xlUp).Row
    .Cells(sr, col).Resize(lr, xrange).ClearContents

End With

End Sub

Public Function BubbleSort_array(arr As Variant, Optional ByVal method As String = "Ascending")

'https://officeguide.cc/excel-vba-array-bubble-sort-function-tutorial-examples/
'https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
'冒泡排序僅供少數資料使用(data<70)

    idxMin = LBound(arr)
    idxMax = UBound(arr)
  
    For i = idxMin To idxMax - 1
      For j = i + 1 To idxMax
      
        If method = "Ascending" Then '升排序
      
            If arr(i) > arr(j) Then ' 判斷條件
              temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        
        Else
        
            If arr(i) < arr(j) Then ' 判斷條件
              temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        
        End If
        
      Next j
    Next i
    
    BubbleSort_array = arr

End Function

Public Function BubbleSort_coll(ByVal coll As Collection, Optional ByVal method As String = "Ascending")

'rely on=>>
'   tranColl2Array
'   tranArray2Coll

    arr = tranColl2Array(coll)
    arr_sort = BubbleSort_array(arr, method)
    
    Set BubbleSort_coll = tranArray2Coll(arr_sort)

End Function

Public Function tranColls2Array(ParamArray colls() As Variant) ''ByVal colls As Collection)

'每一集合儲存成一維度

Dim arr()
Dim k As Integer

ReDim arr(0 To colls(0).Count - 1, UBound(colls))

m = 0

For Each coll In colls

    k = 0
    
    For Each it In coll
    
        arr(k, m) = it
        k = k + 1
    
    Next

    m = m + 1
    
Next

tranColls2Array = arr

End Function

Public Function tranColls2OneArray(ParamArray colls() As Variant) As Double()

'各集合依序儲存成同一維度 (常見於XY座標集合轉換成Vertices)

Dim arr() As Double
ReDim arr(0 To colls(0).Count * (UBound(colls) + 1) - 1)

For i = 1 To colls(0).Count

    For j = 0 To UBound(colls)
    
        arr(k) = colls(j)(i)
        k = k + 1
    
    Next

Next

tranColls2OneArray = arr

End Function

Public Function tranColl2Array(ByVal coll As Collection, Optional lboundcount As Integer = 0) As Variant

Dim arr()
Dim k As Integer

ReDim arr(0 + lboundcount To coll.Count - 1 + lboundcount)

For i = 1 To coll.Count
    arr(0 + lboundcount + k) = coll(i)
    k = k + 1
Next

tranColl2Array = arr

End Function

Public Function tranArray2Coll(ByVal arr As Variant) As Collection

Dim coll As New Collection

For i = LBound(arr) To UBound(arr)
    coll.Add arr(i)
Next

Set tranArray2Coll = coll

End Function

Public Sub showList(ByVal Data As Variant)

cnt = 1
If Not TypeName(Data) = "Collection" Then cnt = LBound(Data)

For Each it In Data
     Debug.Print cnt & " => " & it: cnt = cnt + 1
Next

End Sub

Function MergeSort_array(ByRef arr As Variant, Optional ByRef sortKeys As Variant = Null, Optional ByVal hasHeaders As Boolean = False)
'Sorts a one or two dimensional array.
'2 dimensional arrays can have their sort keys specified by passing
'the appropriate column number(s) as the sortKeys parameter.
'Function passes a reference so will mutate your original array.
'If this is not desirable you must pass a copy.
'
'Example uses:
' sortArray myArray - One-dimensional array
' sortArray myArray, 2 - Two-dimensional array, single sort key
' sortArray myArray, Array(2,3,1) - Two-dimensional array, multiple sort keys
' sortArray myArray, Array(2,3,1), True - Two-dimensional array, multiple sort keys with headers preserved

'https://www.ozgrid.com/forum/index.php?thread/1220720-vba-merge-sort/

Dim mid As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim X As Long
Dim Y As Long
Dim sortMode As Long
Dim arr1
Dim arr2
Dim head
Dim tmp

If UBound(arr) - LBound(arr) = 0 Then Exit Function

On Error Resume Next

i = UBound(arr, 2)
If Err.Number <> 0 Then
    sortMode = 1 'Not a 2D array
    If hasHeaders Then
        ReDim tmp(LBound(arr) To UBound(arr) - 1)
        ReDim head(1 To 1)
        For i = LBound(arr) To UBound(arr)
            If i = LBound(arr) Then
                head(1) = arr(LBound(arr))
            Else
                tmp(i - 1) = arr(i)
            End If
            
        Next i
        arr = tmp
    End If
Else
    sortMode = 2
    If hasHeaders Then
        ReDim tmp(LBound(arr) To (UBound(arr) - 1), LBound(arr, 2) To UBound(arr, 2))
        ReDim head(1 To 1, LBound(arr, 2) To UBound(arr, 2))
        For i = LBound(arr) To UBound(arr)
            For j = LBound(arr, 2) To UBound(arr, 2)
                If i = LBound(arr) Then
                    head(1, j) = arr(LBound(arr), j)
                Else
                    tmp(i - 1, j) = arr(i, j)
                End If
            Next j
        Next i
        arr = tmp
    End If
End If

On Error GoTo 0

If IsNumeric(sortKeys) Then
    sortKeys = Array(CLng(sortKeys))
ElseIf IsNull(sortKeys) Then
    sortKeys = Array(LBound(arr))
End If

Y = LBound(sortKeys)
mid = Int((UBound(arr) + IIf(LBound(arr) = 0, 1, 0)) / 2)

If mid < LBound(arr) Then mid = LBound(arr)

If sortMode = 1 Then

    ReDim arr1(LBound(arr) To mid - IIf(LBound(arr) = 0, 1, 0))
    ReDim arr2(LBound(arr) To UBound(arr) - mid)
    j = LBound(arr)
    For i = LBound(arr1) To UBound(arr1)
        arr1(i) = arr(j)
        j = j + 1
    Next i
    For i = LBound(arr2) To UBound(arr2)
        arr2(i) = arr(j)
        j = j + 1
    Next i
ElseIf sortMode = 2 Then

    ReDim arr1(LBound(arr) To mid - IIf(LBound(arr) = 0, 1, 0), LBound(arr, 2) To UBound(arr, 2))
    ReDim arr2(LBound(arr) To UBound(arr) - mid, LBound(arr, 2) To UBound(arr, 2))
    j = LBound(arr)
    For i = LBound(arr1) To UBound(arr1)
        For k = LBound(arr1, 2) To UBound(arr1, 2)
            arr1(i, k) = arr(j, k)
        Next k
        j = j + 1
    Next i
    
    For i = LBound(arr2) To UBound(arr2)
        For k = LBound(arr2, 2) To UBound(arr2, 2)
            arr2(i, k) = arr(j, k)
        Next k
     j = j + 1
    Next i
End If

MergeSort_array arr1, sortKeys 'I call myself!!!
MergeSort_array arr2, sortKeys 'I call myself again!!!
i = LBound(arr)
j = LBound(arr1)
k = LBound(arr2)

If sortMode = 1 Then
    While j <= UBound(arr1) And k <= UBound(arr2)
        If arr1(j) <= arr2(k) Then
        arr(i) = arr1(j)
        j = j + 1
        Else
        arr(i) = arr2(k)
        k = k + 1
        End If
        i = i + 1
    Wend
    While j <= UBound(arr1)
        arr(i) = arr1(j)
        j = j + 1
        i = i + 1
    Wend
    While k <= UBound(arr2)
        arr(i) = arr2(k)
        k = k + 1
        i = i + 1
    Wend
ElseIf sortMode = 2 Then
While j <= UBound(arr1) And k <= UBound(arr2)
    If arr1(j, sortKeys(Y)) < arr2(k, sortKeys(Y)) _
    Or (arr1(j, sortKeys(Y)) = arr2(k, sortKeys(Y)) And UBound(sortKeys) = Y) Then
    For X = LBound(arr1, 2) To UBound(arr1, 2)
        arr(i, X) = arr1(j, X)
        Next X
        j = j + 1
        Y = LBound(sortKeys)
    ElseIf arr1(j, sortKeys(Y)) > arr2(k, sortKeys(Y)) Then
        For X = LBound(arr2, 2) To UBound(arr2, 2)
            arr(i, X) = arr2(k, X)
        Next X
        k = k + 1
        Y = LBound(sortKeys)
    Else
        i = i - 1
         Y = Y + 1
    End If
    i = i + 1
Wend
While j <= UBound(arr1)
    For X = LBound(arr1, 2) To UBound(arr1, 2)
    arr(i, X) = arr1(j, X)
    Next X
    j = j + 1
    i = i + 1
Wend
While k <= UBound(arr2)
    For X = LBound(arr2, 2) To UBound(arr2, 2)
    arr(i, X) = arr2(k, X)
    Next X
    k = k + 1
    i = i + 1
Wend
End If
If hasHeaders Then
    If sortMode = 1 Then
        '1d
        ReDim tmp(LBound(tmp) To UBound(tmp) + 1)
        tmp(LBound(tmp)) = head(1)
        For i = LBound(arr) To UBound(arr)
            tmp(i + 1) = arr(i)
        Next i
    Else
        '2d
        ReDim tmp(LBound(tmp) To UBound(tmp) + 1, LBound(tmp, 2) To UBound(tmp, 2))
        For i = LBound(tmp) To UBound(tmp)
            For j = LBound(tmp, 2) To UBound(tmp, 2)
            If i = LBound(tmp) Then
                tmp(i, j) = head(1, j)
            Else
                tmp(i, j) = arr(i - 1, j)
            End If
            Next
        Next i
    End If
    arr = tmp
End If
End Function

Public Function ReverseColl(ByVal coll)

Dim coll_tmp As New Collection

For i = coll.Count To 1 Step -1

    coll_tmp.Add coll(i)

Next

Set ReverseColl = coll_tmp

End Function

Public Function combineArray(arr1 As Variant, arr2 As Variant)

Dim arr()

ReDim arr(LBound(arr1) To UBound(arr1) + UBound(arr2) + 1)

For Each a In arr1
    
    arr(cnt) = a
    cnt = cnt + 1
Next

For Each a In arr2

    arr(cnt) = a
    cnt = cnt + 1

Next

combineArray = arr

End Function

Function SortPTArray(ByVal arr)

Dim collX As New Collection
Dim collY As New Collection

For i = LBound(arr) To UBound(arr) Step 3

    collX.Add arr(i)
    collY.Add arr(i + 1)

Next

SortPTArray = tranColls2Array(collX, collY)

Call MergeSort_array(SortPTArray)

End Function


'Public Function getItemFromPTArray(ByVal arr, ByVal j As Byte, ByVal mode As String)
'
'Dim coll As New Collection
'
'For i = LBound(arr) To UBound(arr) Step 3
'
'    coll.Add arr(i + j)
'
'Next
'
'Set sort_coll = BubbleSort_coll(coll)
'
'If mode = "MAX" Then
'    getItemFromPTArray = sort_coll(sort_coll.Count)
'Else
'    getItemFromPTArray = sort_coll(1)
'End If
'
'End Function


