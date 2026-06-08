Function CountFilesInFolder(folderPath As String) As Long
    Dim fso As Object
    Dim targetFolder As Object

    On Error GoTo ErrHandler

    folderPath = Trim(folderPath)

    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)

    CountFilesInFolder = targetFolder.Files.Count
    Exit Function

ErrHandler:
    CountFilesInFolder = -1
End Function