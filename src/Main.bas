Attribute VB_Name = "Main"
Option Explicit

Public swApp As Object
Public gFSO As FileSystemObject
Dim gDocName As String
Dim gDrawingName As String

Sub Main()

  Dim CurrentDoc As ModelDoc2
  Dim OpenedWindow As ModelWindow
  Dim IsNewDoc As Boolean
  Dim UserChoice As VbMsgBoxResult

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
    gDrawingName = CreateDrawingName(gDocName)
    SearchOpenedDrawing gDrawingName, OpenedWindow
    If There(OpenedWindow) Then
      OpenedWindow.Activate
    ElseIf gFSO.FileExists(gDrawingName) Then
      AskOkCancel "Открыть существующий чертеж?", UserChoice
      If UserChoice = vbOK Then
        OpenDrawing gDrawingName
      Else
        ChooseTemplateAndCreateDrawing
      End If
    Else
      ChooseTemplateAndCreateDrawing
    End If
  End If
 
End Sub

Function Run() 'hide

  Dim Template As String
  
  Template = ChooseTemplateDialog.ListBoxNames.Text
  ChooseTemplateDialog.Hide
  CreateDrawing Template, gDocName, gDrawingName
  ExitApp

End Function

Function ExitApp() 'hide

  Unload ChooseTemplateDialog
  End

End Function

Function SearchFormInit() 'hide

  Dim i As Variant
  
  For Each i In GetDrawingTemplates
    ChooseTemplateDialog.ListBoxNames.AddItem i
  Next
  If ChooseTemplateDialog.ListBoxNames.ListCount > 0 Then
    ChooseTemplateDialog.ListBoxNames.ListIndex = 0
  End If

End Function

Function ChooseTemplateAndCreateDrawing() 'hide

  ChooseTemplateDialog.Show
  
End Function
