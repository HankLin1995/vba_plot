Attribute VB_Name = "GIT"
Sub ExportCodesToFolder()

'Type: 1=bas,2=cls,3=frm

myFolder = getSavedFolder
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

'MsgBox "�Ъ`�N: doccls���ɦW�ثe�ٵL�k�i��פJ�A�ݭn��ʳB�z!", vbInformation
    
End Sub

Sub ImportCodes()

Set coll_path = GetFilePathsInFolder

For Each filePath In coll_path

    Filename = mid(filePath, InStrRev(filePath, "\") + 1)
    fileExtension = mid(Filename, InStrRev(Filename, ".") + 1)
    'fileName_short = mid(Filename, 1, InStrRev(Filename, ".") - 1)
    
    If fileExtension = "frm" Or fileExtension = "bas" Or fileExtension = "cls" Then
        Call ImportCode(filePath, Filename)
    End If

Next

End Sub

Sub ImportCode(ByVal filePath As String, ByVal Filename As String)

extension = mid(Filename, InStrRev(Filename, ".") + 1)
CodeName = mid(Filename, 1, InStrRev(Filename, ".") - 1)

Set VBProj = ThisWorkbook.VBProject

If checkIfCodeExist(CodeName) = True Then

    Set VBComp = VBProj.VBComponents(CodeName)
    VBProj.VBComponents.Remove (VBComp)

End If

VBProj.VBComponents.Import (filePath)

End Sub

'--------FUNCTION------------

Function GetFilePathsInFolder()

    Dim coll As New Collection

    Dim fso As Object
    Dim folderPath As String
    Dim folder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    folderPath = getSavedFolder
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

    Set fldr = Application.FileDialog(4) 'msoFileDialogFolderPicker
    
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then FolderName = .SelectedItems(1)
    End With
getSavedFolder = FolderName

End Function

Function checkIfCodeExist(ByVal checkName As String)

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


