VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CategoriesQuickSelector 
   Caption         =   "Quick Select Categories"
   ClientHeight    =   9495
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   13500
   OleObjectBlob   =   "CategoriesQuickSelector.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "CategoriesQuickSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub setCategories()
Dim Sel As Outlook.Selection

End Sub



Private Sub CommandButton1_Click()
' Depricated
Dim obj As Object
Dim Sel As Outlook.Selection
Dim list As ListBox

Set list = CategoriesQuickSelector.CategoriesList
Set Sel = Application.ActiveExplorer.Selection


If Sel.Count = 0 Then
      Exit Sub
Else
      Set obj = Sel(1)
End If

Debug.Print (obj.Categories)


For i = 0 To (list.ListCount - 1)
    If list.Selected(i) Then
        obj.Categories = obj.Categories & ";" & list.list(i)
        obj.Save
    End If
Next i
Call Reset_categories_list

End Sub



Private Sub Add_Click()
    If Application.ActiveExplorer.Selection.Count = 0 Then
        Exit Sub
'    ElseIf Application.ActiveExplorer.Selection.Count = 1 Then
'        Call Add_Click_single
    ElseIf Application.ActiveExplorer.Selection.Count > 0 Then
        Call Add_Click_multi
    End If


End Sub


Private Sub Add_Click_single()
Dim obj As Object
Dim Sel As Outlook.Selection
Dim list As ListBox

Set list = CategoriesQuickSelector.CategoriesList
Set Sel = Application.ActiveExplorer.Selection


If Sel.Count = 0 Then
      Exit Sub
Else
      Set obj = Sel(1)
End If

Debug.Print (obj.Categories)


For i = 0 To (list.ListCount - 1)
    If list.Selected(i) Then
        obj.Categories = obj.Categories & ";" & list.list(i)
        obj.Save
    End If
Next i

Call Reset_categories_list


End Sub

Private Sub Add_Click_multi2()


Dim obj As Object
Dim Sel As Outlook.Selection
Dim list As ListBox
Dim Tags As ListBox

Dim existingCats As String

Set list = CategoriesQuickSelector.CategoriesList
Set Tags = CategoriesQuickSelector.Taglist
Set Sel = Application.ActiveExplorer.Selection


If Sel.Count = 0 Then
      Exit Sub
End If


For Each obj In Sel

    existingCats = obj.Categories
    
    obj.Categories = ""
    
    For i = 0 To (list.ListCount - 1)
        For Each X In Split(existingCats, ";")
            If list.Selected(i) Then
                If Not (StrComp(X, list.list(i)) = 0) Then
                    obj.Categories = obj.Categories & ";" & list.list(i)
                    obj.Save
                End If
            End If
        Next X
    Next i

Next obj

Call Reset_categories_list



End Sub

Private Sub Add_Click_multi()
Dim obj As Object
Dim Sel As Outlook.Selection
Dim list As ListBox
Dim Tags As ListBox

Dim tobedeleted As Boolean

Dim existingCats As Variant

Dim xAppointmentItem As Outlook.AppointmentItem


Set list = CategoriesQuickSelector.CategoriesList
Set Tags = CategoriesQuickSelector.Taglist
Set Sel = Application.ActiveExplorer.Selection


If Sel.Count = 0 Then
      Exit Sub
End If


For Each obj In Sel

    'existingCats = Split(obj.Categories, ";")
    
    'obj.Categories = ""
    'obj.Save
    
    
    
        For i = 0 To (list.ListCount - 1)
            If list.Selected(i) Then
                If Not (ExistsIn(obj.Categories, list.list(i))) Then
                    'If TypeOf obj Is MeetingItem Then
                    '    obj.GetAssociatedAppointment(True).Categories = obj.GetAssociatedAppointment(True).Categories & ";" & list.list(i)
                    'Else
                        obj.Categories = obj.Categories & ";" & list.list(i)
                    'End If
                    obj.Save
                End If
            End If
        Next i

Next obj


CategoriesQuickSelector.TextBox1.text = ""
Call TextBox1_Change
End Sub


Private Sub CategoriesList_AfterUpdate()
With CategoriesQuickSelector.CategoriesList
For i = 0 To .ListCount - 1
    .Selected(i) = False
Next
End With
End Sub

Private Sub CategoriesList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call Add_Click
End Sub

Private Sub Taglist_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call Remove_Click
    
End Sub

Private Sub Remove_Click()
    If Application.ActiveExplorer.Selection.Count = 0 Then
        Exit Sub
'    ElseIf Application.ActiveExplorer.Selection.Count = 1 Then
'        Call Remove_Click_single
    ElseIf Application.ActiveExplorer.Selection.Count > 0 Then
        Call Remove_Click_multi
    End If
        
End Sub


Private Sub Remove_Click_single()
Dim obj As Object
Dim Sel As Outlook.Selection
Dim list As ListBox
Dim Tags As ListBox

Set list = CategoriesQuickSelector.CategoriesList
Set Tags = CategoriesQuickSelector.Taglist
Set Sel = Application.ActiveExplorer.Selection


If Sel.Count = 0 Then
      Exit Sub
Else
      Set obj = Sel(1)
End If

obj.Categories = ""

For i = 0 To (Tags.ListCount - 1)
    If Not Tags.Selected(i) Then

            obj.Categories = obj.Categories & ";" & Tags.list(i)
            obj.Save

    End If
Next i

Call Reset_categories_list
End Sub




Private Sub Remove_Click_multi()
Dim obj As Object
Dim Sel As Outlook.Selection
Dim list As ListBox
Dim Tags As ListBox

Dim tobedeleted As Boolean

Dim existingCats As Variant

Set list = CategoriesQuickSelector.CategoriesList
Set Tags = CategoriesQuickSelector.Taglist
Set Sel = Application.ActiveExplorer.Selection


If Sel.Count = 0 Then
      Exit Sub
End If


For Each obj In Sel

    existingCats = Split(obj.Categories, ";")
    
    obj.Categories = ""
    obj.Save
    
    
    For Each X In existingCats
        tobedeleted = False
        For i = 0 To (Tags.ListCount - 1)
            If Tags.Selected(i) Then
                If (StrComp(Trim(X), Tags.list(i)) = 0) Then
                    tobedeleted = True
                    Exit For
                End If
            End If
        Next i
        
        If tobedeleted Then
            X = ""
        End If
        
        If Not (tobedeleted) Then
            obj.Categories = obj.Categories & ";" & Trim(X)
            obj.Save
        End If
        
    Next X

Next obj

Call Reset_categories_list
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





Private Sub TextBox1_Change()
Dim text As String
Dim list As ListBox
Dim addit As Boolean



Set list = CategoriesQuickSelector.CategoriesList

text = CategoriesQuickSelector.TextBox1.text

If Not text = "" Then

    list.Clear
    For Each cat In Application.Session.Categories
        If Not (InStr(1, cat, text, vbTextCompare) = 0) Then
                    addit = True
                    For Each X In Split(Application.ActiveExplorer.Selection(1).Categories, ";")
                        If StrComp(cat, Trim(X)) = 0 Then
                            addit = False
                            Exit For
                        Else
                            addit = True
                        End If
                    Next X
                    If addit Then
                        list.AddItem (cat)
                    End If
            
        End If
    Next cat

    With list
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.list(i)) > LCase(.list(i + 1)) Then
                    temp = .list(i)
                    .list(i) = .list(i + 1)
                    .list(i + 1) = temp
                End If
            Next i
        Next j
    End With

Else
    Call Reset_categories_list
End If


End Sub

Private Sub Tags_Update()

Dim Tags As ListBox

Set list = CategoriesQuickSelector.CategoriesList
Set Tags = CategoriesQuickSelector.Taglist

    For i = 0 To (list.ListCount - 1)
        For Each X In Split(Application.ActiveExplorer.Selection(1).Categories, ";")
            If StrComp(list.list(i), Trim(X)) = 0 Then
                'Debug.Print (Trim(x))
                Debug.Print (list.list(i))
                Debug.Print (Trim(X))
                Debug.Print (StrComp(list.list(i), Trim(X)))
                Tags.AddItem (list.list(i))
                list.RemoveItem (i)
            End If
        Next X
    Next i
    
    With Tags
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.list(i)) > LCase(.list(i + 1)) Then
                    temp = .list(i)
                    .list(i) = .list(i + 1)
                    .list(i + 1) = temp
                End If
            Next i
        Next j
    End With

End Sub

Private Sub Reset_categories_list()
Dim Categories As Outlook.Categories

Dim list As ListBox
Dim Tags As ListBox
Dim addit As Boolean
Dim cat As Category
Dim i As Long

Dim item As mailitem

Dim usedcats() As String

addit = False
Set list = CategoriesQuickSelector.CategoriesList
Set Tags = CategoriesQuickSelector.Taglist

list.Clear
Tags.Clear

If Not Application.ActiveExplorer.Selection.Count = 0 Then


For Each cat In Application.Session.Categories
    If Application.ActiveExplorer.Selection(1).Categories = "" Then
        list.AddItem (cat)
    Else
        addit = False
        For Each X In Split(Application.ActiveExplorer.Selection(1).Categories, ";")
            If StrComp(cat, Trim(X)) = 0 Then

                Tags.AddItem (cat)
                
                addit = False
                Exit For
            Else
                addit = True
            End If
        Next X
        If addit Then
            list.AddItem (cat)
        End If
    End If
Next cat

End If
    
With list
    For j = 0 To .ListCount - 2
        For i = 0 To .ListCount - 2
            If LCase(.list(i)) > LCase(.list(i + 1)) Then
                temp = .list(i)
                .list(i) = .list(i + 1)
                .list(i + 1) = temp
            End If
        Next i
    Next j
End With
    
If Not CategoriesQuickSelector.TextBox1.text = "" Then
    Call TextBox1_Change
End If




With CategoriesQuickSelector.Taglist
    For i = 0 To .ListCount - 1
        .Selected(i) = False
    Next i
End With
    
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
         
        Dim list As ListBox
        Set list = CategoriesQuickSelector.CategoriesList
        If list.ListCount = 1 Then
            list.Selected(0) = True
            Call Add_Click
            CategoriesQuickSelector.TextBox1.Value = ""
            Call TextBox1_Change
        End If
         'Whatever other code that does not affect the focus again...
    End If
End Sub
Private Sub update_Click()
    
    Call Reset_categories_list
    
End Sub


Public Sub UserForm_Initialize()
    
    Call Reset_categories_list
    
End Sub '
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Call Reset_categories_list
    
End Sub '
