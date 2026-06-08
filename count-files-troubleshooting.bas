Sub ListFilesInFolderDebug()
    Dim fso As Object
    Dim targetFolder As Object
    Dim f As Object
    Dim folderPath As String

    folderPath = "\\rspnc2data\public\SAP B1 - Public\Test Scripts March 2026\Synoptics Test Scripts UAT #2\Accounts Payable\Not Relevant"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = fso.GetFolder(folderPath)

    For Each f In targetFolder.Files
        Debug.Print f.Name, f.Attributes
    Next f
End Sub