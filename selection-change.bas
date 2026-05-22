Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error GoTo SafeExit

    'Only trigger for single-cell selections in column A
    If Target.CountLarge > 1 Then Exit Sub
    If Target.Column <> 1 Then Exit Sub
    If Target.Row < 5 Then Exit Sub
    If Trim(Target.Value) = "" Then Exit Sub

    Application.EnableEvents = False

    '''DrawBouleGradient - ### Change this to function that needs to be automated'''

SafeExit:
    Application.EnableEvents = True

End Sub
