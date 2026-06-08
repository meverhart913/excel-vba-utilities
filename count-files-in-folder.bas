Function CountVisibleFilesInFolder(folderPath As String) As Long
    Dim fso As Object
    Dim targetFolder As Object
    Dim f As Object
    Dim fileCount As Long

    Application.Volatile

    On Error GoTo ErrHandler

    folderPath = Trim(folderPath)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)

    For Each f In targetFolder.Files
        Debug.Print f.Name & " | Attr=" & f.Attributes & " | HiddenTest=" & (f.Attributes And 2)

        If (f.Attributes And 2) <> 2 Then
            fileCount = fileCount + 1
        End If
    Next f

    CountVisibleFilesInFolder = fileCount
    Exit Function

ErrHandler:
    CountVisibleFilesInFolder = -1
End Function