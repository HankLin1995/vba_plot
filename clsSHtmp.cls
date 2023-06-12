VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSHtmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtSH As Object
Private shtH As Object

Public myWidth As Double

Const LEFT = "左支距"
Const RIGHT = "右支距"

Private Sub Class_Initialize()

Set shtSH = Sheets("標竿橫斷TMP")
Set shtH = Sheets("橫斷面")

End Sub

Sub ChangeLoc()

With Sheets("橫斷面")

lr = .Cells(.Rows.Count, 4).End(xlUp).Row

For r = 2 To lr

    If .Cells(r, 4) <> "" Then
    
        If .Cells(r, 4) Like "*.*" Then
        
            tmp = Split(.Cells(r, 4), ".")
        
            If tmp(1) = "1" Then
                s = Format(tmp(0), "0+000") & "(上)"
            Else
                s = Format(tmp(0), "0+000") & "(下)"
            End If
        
        Else
            
            s = Format(.Cells(r, 4), "0+000")
        
        End If
    
        .Cells(r, 4) = s
    
    End If

Next

End With

End Sub


Sub test() '就是個測試

arr1 = Array(0.15, 0.45, 0.4, 0)
arr2 = Array(0, 0.45, 0, -0.1)
arr3 = Array("RWC", "RW", "RT", "FA")

For i = 0 To UBound(arr1)

    SHTable.txtX = arr1(i)
    SHTable.txtY = arr2(i)
    SHTable.txtCD = arr3(i)
    
    Call SHTable.cmdAdd_Click

Next

SHTable.txtCLCD = "CL"
SHTable.txtLoc = "0+000"
SHTable.txtCLHeight = 50

End Sub

Sub ImportData() '匯入表單資料

Dim coll As New Collection

With SHTable

    myLocation = .txtLoc
    myCLCD = .txtCLCD
    myCLH = .txtCLHeight
    myDir = .OptionButton2.Value

End With

With SHTable.lstTMP

    For i = 0 To .ListCount - 1
        
        s = s & "," & .List(i, 0) & "," & .List(i, 1) & "," & .List(i, 2)
    
    Next

End With

With shtSH

    IsDataExist = checkData(myLocation, myDir)

    If IsDataExist Then MsgBox "已經填寫過了!": Exit Sub

    lr = .Cells(Rows.Count, 1).End(xlUp).Row + 1

    .Cells(lr, 1) = myLocation
    .Cells(lr, 2) = myCLCD
    .Cells(lr, 3) = myCLH
    .Cells(lr, 4) = LEFT

    If myDir = True Then .Cells(lr, 4) = RIGHT
    
    .Cells(lr, 5) = mid(s, 2)

End With

End Sub

Private Function checkData(ByVal myLocation As String, ByVal myDir As Boolean) As Boolean

IsDataExist = False

If myDir = True Then
    myDirCheck = RIGHT
Else
    myDirCheck = LEFT
End If

With shtSH

    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        If .Cells(r, 1) = myLocation And .Cells(r, 4) = myDirCheck Then
        
            checkData = True
            
        End If
    Next

End With

End Function

Sub Export()

Application.EnableEvents = False

With shtSH
    
    '-----Init-----
    
    sr = 3
    er = .Cells(Rows.Count, 1).End(xlUp).Row
    
    .Activate
    .Range("A" & sr & ":E" & er).Sort key1:=.Range("D" & sr & ":D" & er), order1:=xlAscending
    .Range("A" & sr & ":E" & er).Sort key1:=.Range("A" & sr & ":A" & er), order1:=xlAscending

    '-----Start------

    For r = sr To er 'Step 2
     
        myloc = .Cells(r, 1)
        nextloc = .Cells(r + 1, 1)
        
        If myloc = nextloc Then
        
            mySHTable = retSHTable(r, 1)
            r = r + 1
        
        Else
        
            mySHTable = retSHTable(r, 0)
        
        End If
        
        Call SHTable2Hsection(mySHTable, myloc)
            
        
    Next

End With

shtH.Activate

Application.EnableEvents = True

End Sub

Sub SHTable2Hsection(ByVal ShapeArr, ByVal myloc As String)

With shtH

lr = .Cells(Rows.Count, 1).End(xlUp).Row + 2

If lr = 3 Then lr = 2

.Cells(lr, 4) = myloc

For i = 1 To UBound(ShapeArr, 1)
    
    For c = 1 To 3
    
        .Cells(lr + j, c) = ShapeArr(i, c)
    
    Next
    
    j = j + 1

Next

End With

End Sub

Function retSHTable(ByVal sr As Integer, ByVal cnt) '結合整個橫斷面

Dim FinalShapeArr As Variant

Dim collSX As New Collection
Dim collSY As New Collection
Dim collSCD As New Collection

With shtSH

    For r = sr To sr + cnt
    
        myLocation = .Cells(r, 1)
        myCLCD = .Cells(r, 2)
        myCLH = .Cells(r, 3)
        myDir = .Cells(r, 4)
        
        tmpArr = retShapeArr(r) '取半邊的橫斷面
        
        collSX.Add 0
        collSY.Add myCLH
        collSCD.Add myCLCD
        
        If myDir = LEFT Then
            
            For i = 1 To UBound(tmpArr)
            
                collSX.Add tmpArr(i, 1), before:=1
                collSY.Add tmpArr(i, 2), before:=1
                collSCD.Add tmpArr(i, 3), before:=1
                
            Next
            
        Else
'
'            collSX.Add 0
'            collSY.Add myCLH
'            collSCD.Add myCLCD
            
            For i = 1 To UBound(tmpArr)
            
                collSX.Add tmpArr(i, 1)
                collSY.Add tmpArr(i, 2)
                collSCD.Add tmpArr(i, 3)
                
            Next
        
        End If
        
    Next
    
    ReDim FinalShapeArr(1 To collSX.Count, 1 To 3)
    
    For i = 1 To collSX.Count
        
        FinalShapeArr(i, 1) = collSX(i) + myWidth
        FinalShapeArr(i, 2) = collSY(i)
        FinalShapeArr(i, 3) = collSCD(i)
        
    Next

retSHTable = FinalShapeArr
    
End With

End Function

Private Function retShapeArr(ByVal r As Integer) '取得半邊的橫斷

Dim ShapeArr As Variant

Dim sumX As Double
Dim sumY As Double

Dim collX As New Collection
Dim collY As New Collection
Dim collCD As New Collection

With shtSH

    myCLH = .Cells(r, 3)
    myDir = .Cells(r, 4)
    myCont = .Cells(r, 5)

End With

tmp = Split(myCont, ",")

For i = 0 To UBound(tmp) Step 3

    X = tmp(i)
    Y = tmp(i + 1)
    CD = tmp(i + 2)
    
    If myDir = LEFT Then
        sumX = sumX - X
    Else
        sumX = sumX + X
    End If
    
    sumY = sumY + Y
    
    collX.Add sumX
    collY.Add sumY + myCLH
    collCD.Add CD
    
Next

ReDim ShapeArr(1 To collX.Count, 1 To 3)

For i = 1 To collX.Count

    ShapeArr(i, 1) = collX(i)
    ShapeArr(i, 2) = collY(i)
    ShapeArr(i, 3) = collCD(i)

Next

retShapeArr = ShapeArr

End Function
