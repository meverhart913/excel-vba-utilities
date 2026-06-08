Function CountFilesInFolder(folderPath As String) As Long
    Dim fso As Object
    Dim targetFolder As Object
    Dim f As Object
    Dim fileCount As Long

    On Error GoTo ErrHandler

    folderPath = Trim(folderPath)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)

    For Each f In targetFolder.Files
        If (f.Attributes And 2) = 0 Then
            fileCount = fileCount + 1
        End If
    Next f

    CountFilesInFolder = fileCount
    Exit Function

ErrHandler:
    CountFilesInFolder = -1
End Function