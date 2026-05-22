Function CountFilesInFolder(folderPath As String) As Long
    Dim fileName As String
    Dim fileCount As Long
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = Dir(folderPath & "*.*")
    
    Do While fileName <> ""
        fileCount = fileCount + 1
        fileName = Dir
    Loop
    
    CountFilesInFolder = fileCount
End Function
