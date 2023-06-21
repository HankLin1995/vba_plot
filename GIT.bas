Attribute VB_Name = "GIT"
'TODO:Export folder need to be killed

Sub ExportCodesToFolder()

'Type: 1=bas,2=cls,3=frm

myFolder = getSavedFolder

Call killFilesInFolder(myFolder)

Set VBProj = ThisWorkbook.VBProject
For Each VBComp In VBProj.VBComponents
    
    Select Case VBComp.Type
    
        Case 1: myExtension = ".bas"
        Case 2: myExtension = ".cls"
        Case 3: myExtension = ".frm"
        
        Case 100: myExtension = ".doccls"
    
    End Select
    
    full_path = myFolder & "\" & VBComp.Name & myExtension
    
    If myExtension <> "" Then
    
        VBComp.Export (full_path)
    
    End If
    
    If myExtension = ".doccls" And CountFileLines(full_path) = 9 Then Kill full_path
    
Next VBComp
    
End Sub

Sub killFilesInFolder(folderPath)

Set coll_path = GetFilePathsInFolder(folderPath)

For Each filePath In coll_path

    Filename = mid(filePath, InStrRev(filePath, "\") + 1)
    fileExtension = mid(Filename, InStrRev(Filename, ".") + 1)
    
    If fileExtension = "frm" Or fileExtension = "bas" Or fileExtension = "cls" Or fileExtension = "doccls" Then
        Kill filePath
    End If
Next

End Sub

Sub ImportCodes()

myFolder = getSavedFolder

Set coll_path = GetFilePathsInFolder(myFolder)

Call DeleteCodes

For Each filePath In coll_path

    Filename = mid(filePath, InStrRev(filePath, "\") + 1)
    fileExtension = mid(Filename, InStrRev(Filename, ".") + 1)
    
    If fileExtension = "frm" Or fileExtension = "bas" Or fileExtension = "cls" Then
        Call ImportCode(filePath, Filename)
    End If

Next

End Sub

Sub ImportCode(ByVal filePath As String, ByVal Filename As String)

extension = mid(Filename, InStrRev(Filename, ".") + 1)
CodeName = mid(Filename, 1, InStrRev(Filename, ".") - 1)

If CodeName = "GIT" Then Exit Sub

Set VBProj = ThisWorkbook.VBProject

'If checkIfCodeExist(CodeName) = True Then
'
'    Set vbcomp = VBProj.VBComponents(CodeName)
'    VBProj.VBComponents.Remove (vbcomp)
'
'End If

VBProj.VBComponents.Import (filePath)

End Sub

Sub DeleteCodes()

'Type: 1=bas,2=cls,3=frm

Set VBProj = ThisWorkbook.VBProject
For Each VBComp In VBProj.VBComponents
    
    Select Case VBComp.Type
    
        Case 1: myExtension = ".bas"
        Case 2: myExtension = ".cls"
        Case 3: myExtension = ".frm"
        
        Case 100: myExtension = ".doccls"
    
    End Select
    
    If VBComp.Type <> 100 And VBComp.Name <> "GIT" Then

        VBProj.VBComponents.Remove (VBComp)
        
    End If
    
Next VBComp

End Sub

'--------FUNCTION------------

Function GetFilePathsInFolder(ByVal folderPath As String)

    Dim coll As New Collection

    Dim fso As Object
    'Dim folderPath As String
    Dim folder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

   ' folderPath = getSavedFolder
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files

        coll.Add file.Path
        
    Next file
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    Set GetFilePathsInFolder = coll
    
End Function

Function getSavedFolder()

    Set fldr = Application.FileDialog(4)
    
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then FolderName = .SelectedItems(1)
    End With
getSavedFolder = FolderName

End Function

Function checkIfCodeExist(ByVal checkName As String) 'useless

Set VBProj = ThisWorkbook.VBProject
Set VBComps = VBProj.VBComponents

checkIfCodeExist = False

For Each it In VBComps

    If it.Name = checkName Then
        
        checkIfCodeExist = True: Exit Function
        
    End If
Next

End Function

Function CountFileLines(ByVal filePath)

    Dim fileContent As String
    Dim fileNumber As Integer
    Dim lineCount As Long
    
    ' Open the text file
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' Read the file content line by line and count the lines
    Do Until EOF(fileNumber)
        Line Input #fileNumber, fileContent
        lineCount = lineCount + 1
    Loop
    
    ' Close the file
    Close fileNumber
    
    ' Display the line count in cell A1
    CountFileLines = lineCount
    
End Function

'--------TMP_CODE-------------

Function tmp_deleteCodes()

Set VBProj = ThisWorkbook.VBProject
Set VBComps = VBProj.VBComponents

For Each it In VBComps
    
    If it.Name Like "*2" And it.Type <> 100 Then

        CodeName = it.Name
        
        Set VBComp = VBProj.VBComponents(CodeName)
        VBProj.VBComponents.Remove (VBComp)
        
    End If
    
Next

End Function



