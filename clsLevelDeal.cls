VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLevelDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtOrg As Worksheet
Private shtExp As Worksheet

Private Sub Class_Initialize()

Set shtOrg = Sheets("測量原始資料")
Set shtExp = Sheets("縱斷面LEVEL")

End Sub

Sub transposeData()

With shtOrg

rr = 1
lr = .Cells(1, 1).End(xlDown).Row
Lc = .Cells(1, 1).End(xlToRight).Column

For r = 2 To lr

    s = .Cells(r, 1)
    
    For c = 2 To Lc
        
        CD = .Cells(1, c)
        H = .Cells(r, c)
        
        If CDbl(H) <> 0 Then
        
            rr = rr + 1
            .Cells(rr, Lc + 2) = s
            .Cells(rr, Lc + 3) = CD
            .Cells(rr, Lc + 4) = H
            .Cells(rr, Lc + 5) = H / 100
        
       ' If r = 16 And c = 5 Then Stop
        
        End If
        
    
    Next

Next

End With

End Sub

Sub collectData()

Dim loc As String
Dim nextloc As String

rr = 2

With shtOrg

    Lc = .Cells(1, 1).End(xlToRight).Column
    lr = .Cells(1, Lc + 2).End(xlDown).Row
    
    shtExp.Cells.Clear
    
    Call InitialTitle(lr)
    
    For r = 2 To lr
        
        loc = .Cells(r, Lc + 2)
        nextloc = .Cells(r + 1, Lc + 2)
        CD = .Cells(r, Lc + 3)
        H = .Cells(r, Lc + 6)
    
        With shtExp
            
            c = getColumn(CD)
        
            If loc Like "*+*" Then
    
                .Cells(rr, 1) = TranLoc(loc) 'Mid(loc, 1, 1) * 1000 + Mid(loc, 3)
                .Cells(rr, c) = H
                
                If nextloc <> loc Then rr = rr + 1
                
            End If
        
        End With
    
    Next

End With

shtExp.Activate

End Sub

Sub InitialTitle(ByVal lr As Integer)

Dim coll_user_cd As New Collection

coll_user_cd.Add "樁號"

With shtOrg

    Set rng = .Rows(1).Find(what:="備註")

    For r = 2 To lr
    
        user_cd = .Cells(r, rng.Column)
        On Error Resume Next: coll_user_cd.Add user_cd, user_cd: On Error GoTo 0
        
    Next
    
    c = 1

    For Each user_cd_item In coll_user_cd
    
        shtExp.Cells(1, c) = user_cd_item
        c = c + 1
    
    Next

End With

End Sub

Sub interpolateData()

With shtExp

lr = .Cells(1, 1).End(xlDown).Row
Lc = .Cells(1, 1).End(xlToRight).Column

    For r = 2 To lr
    
        For c = 2 To Lc
        
            If .Cells(r, c) = 0 Then
            
                ur = .Cells(r, c).End(xlUp).Row
                dr = .Cells(r, c).End(xlDown).Row
                
                If Not ur = 1 And .Cells(dr, c) <> 0 Then
                
                s = interpolate(r, ur, dr, c)
                .Cells(r, c) = Round(s, 2)
                .Cells(r, c).Interior.ColorIndex = 42
                End If
                
            End If
        
        Next
    
    Next

End With

End Sub

Sub ExtractToLSectionSort()

With shtExp

    Lc = .Cells(1, 200).End(xlToLeft).Column
    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    '.Range("A2:" & Chr(Lc + 64) & lr).Sort key1:=.Range("A2:A" & lr), order1:=xlAscending
    .Range("A1:" & Chr(Lc + 64) & lr).Copy
    .Range(Chr(Lc + 65) & "1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    .Columns("A:" & Chr(Lc + 64)).Delete
    .Activate

End With

End Sub

Private Function interpolate(ByVal r, ByVal ur, ByVal dr, ByVal c)

With shtExp

nc = .Cells(r, 1)
sc = .Cells(ur, 1)
ec = .Cells(dr, 1)
sh = .Cells(ur, c)
eH = .Cells(dr, c)

l = ec - sc
interpolate = (eH * (nc - sc) + sh * (ec - nc)) / l

End With

End Function

Private Function getColumn(ByVal CD As String)

With shtExp

    Lc = .Cells(1, 1).End(xlToRight).Column

    For c = 2 To Lc
    
        If .Cells(1, c) = CD Then getColumn = c: Exit Function
    
    Next

End With

End Function


