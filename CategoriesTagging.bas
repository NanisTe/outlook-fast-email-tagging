Attribute VB_Name = "CategoriesTagging"
Sub Tags()

CategoriesQuickSelector.Show

End Sub

Sub TagsAngleichen_addmissingtags()

Dim obj1, obj2 As Object
Dim Sel As Outlook.Selection

Set Sel = Application.ActiveExplorer.Selection


If Sel.Count = 0 Then
      Exit Sub
End If

For Each obj1 In Sel
    For Each obj2 In Sel
        For Each cat In Split(obj1.Categories, ";")
            If Not (ExistsIn(obj2.Categories, Trim(cat))) Then
                Debug.Print (obj2.Subject)
                Debug.Print ("has " & obj1.Categories)
                Debug.Print ("add " & obj2.Categories & ";" & Trim(cat))
                obj2.Categories = obj2.Categories & ";" & cat
                obj2.Save
            End If
        Next cat
    Next obj2
Next obj1

End Sub

Private Function ExistsIn(str_list As String, str As Variant) As Boolean
    Dim exists As Boolean
    exists = False
    For Each X In Split(str_list, ";")
        If StrComp(Trim(X), str) = 0 Then
            exists = True
        End If
    Next X
    ExistsIn = exists
End Function
