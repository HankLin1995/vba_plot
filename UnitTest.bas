Attribute VB_Name = "UnitTest"
Sub getBlank()

Dim myfunc As New clsMyfunction

Set sht = Sheets("¾îÂ_­±")

With sht

    Set coll = myfunc.getBlankColl(sht, 1)
    
    For i = 1 To coll.Count - 1
    
        sr = coll(i) + 1
        er = coll(i + 1) - 1
        
        .Range("A" & sr & ":C" & er).Sort key1:=.Range("A" & sr & ":A" & er), order1:=xlAscending

        Debug.Print .Cells(sr, "D") & ">" & i
    
    Next

End With

End Sub

Function getZfromPT(ByVal strPT As String)
With Sheets("Á`ªí")

    Set rng = .Columns("A").Find(strPT)
    
    If rng Is Nothing Then
        getZfromPT = 0
    Else
        getZfromPT = CDbl(.Cells(rng.Row, "D"))
    End If

End With

End Function

Sub unittest_getZFromPT()

Debug.Assert getZfromPT("33") = 6.897
Debug.Assert getZfromPT("912") = 0

End Sub
