Attribute VB_Name = "Main"
Option Explicit

Const kChooseNewDrawing = "[Создать новый...]"
Public swApp As Object
Public gFSO As FileSystemObject
Dim gDocName As String
Dim gExistDrawings As Dictionary

Sub Main()

  Dim CurrentDoc As ModelDoc2
  Dim IsNewDoc As Boolean
  Dim Name As String
  Dim Designation As String
  Dim FolderPath As String

  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  If CurrentDoc.GetType = swDocDRAWING Then Exit Sub
  
  GetDocNameAndSaveIfNecessary CurrentDoc, IsNewDoc, gDocName
  If gDocName = "" Then
    Exit Sub
  End If
  
  If IsNewDoc Then
    ChooseTemplateAndCreateDrawing
  Else
    SplitNameAndDesignation gFSO.GetBaseName(gDocName), Designation, Name
    FolderPath = gFSO.GetParentFolderName(gDocName)
    Set gExistDrawings = SearchDrawings(Designation, FolderPath)
    If gExistDrawings.Count = 0 Then
      ChooseTemplateAndCreateDrawing
    Else
      ChooseDrawingOrCreateNew
    End If
  End If
  
End Sub

Function RunChooseExistDrawing() 'hide

  Dim OpenedWindow As ModelWindow
  Dim DrawingName As String
  
  With SelectExistDrawingDialog
    .Hide
    If .DrawingListBox.Text = kChooseNewDrawing Then
      ChooseTemplateAndCreateDrawing
    Else
      DrawingName = gExistDrawings(.DrawingListBox.Text)
      SearchOpenedDrawing DrawingName, OpenedWindow
      If There(OpenedWindow) Then
        OpenedWindow.Activate
      Else
        OpenDrawing DrawingName
      End If
    End If
  End With
  ExitApp

End Function

Function RunCreateNewDrawing() 'hide

  Dim Template As String
  
  ChooseTemplateDialog.Hide
  Template = ChooseTemplateDialog.ListBoxNames.Text
  CreateDrawing Template, gDocName, CreateDrawingName(gDocName)
  ExitApp

End Function

Function ExitApp() 'hide

  Unload ChooseTemplateDialog
  Unload SelectExistDrawingDialog
  End

End Function

Function SearchFormInit() 'hide

  Dim I As Variant
  
  With ChooseTemplateDialog.ListBoxNames
    For Each I In GetDrawingTemplates
      .AddItem I
    Next
    If .ListCount > 0 Then
      .ListIndex = 0
    End If
  End With

End Function

Function ChooseTemplateAndCreateDrawing() 'hide

  ChooseTemplateDialog.Show
  
End Function

Function ChooseDrawingFormInit() 'hide

  Dim I As Variant
  
  With SelectExistDrawingDialog.DrawingListBox
    .AddItem kChooseNewDrawing
    For Each I In gExistDrawings.Keys
      .AddItem I
    Next
    If .ListCount > 1 Then
      .ListIndex = 1
    Else
      .ListIndex = 0
    End If
  End With

End Function

Function ChooseDrawingOrCreateNew() 'hide

  SelectExistDrawingDialog.Show

End Function
