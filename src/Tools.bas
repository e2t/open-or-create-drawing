Attribute VB_Name = "Tools"
Option Explicit

Function There(Obj As Object) As Boolean

  There = Not Obj Is Nothing

End Function

Sub AskOkCancel(Prompt As String, ByRef UserChoice As VbMsgBoxResult)

  UserChoice = MsgBox(Prompt, vbOKCancel)

End Sub

Sub CreateDrawing(Template As String, ModelName As String, DrawingName As String)

  Dim NewDoc As ModelDoc2
  Dim Drawing As DrawingDoc
  Dim ASheet As Sheet
  Dim Width As Double
  Dim Height As Double

  Set NewDoc = swApp.NewDocument(Template, swDwgPapersUserDefined, 0, 0)
  Set Drawing = NewDoc
  Set ASheet = Drawing.GetCurrentSheet
  ASheet.GetSize Width, Height
  Drawing.CreateDrawViewFromModelView3 ModelName, "*Front", Width / 2, Height / 2, 0
  NewDoc.SetTitle2 gFSO.GetBaseName(DrawingName)
  swApp.SetCurrentWorkingDirectory gFSO.GetParentFolderName(DrawingName)

End Sub

Function GetDrawingTemplates() As Collection

  Dim i As Variant
  Dim J As Variant
  Dim AFolder As Folder
  Dim AFile As File
  Dim Extension As String
  Dim Templates As Collection
  Dim DefaultTemplate As String
  
  Set Templates = New Collection
  For Each i In GetTemplateLocations
    Set AFolder = gFSO.GetFolder(i)
    For Each J In AFolder.Files
      Set AFile = J
      Extension = gFSO.GetExtensionName(AFile.Name)
      If StrComp(Extension, "DRWDOT", vbTextCompare) = 0 Then
        Templates.Add AFile.Path
      End If
    Next
  Next
  If Templates.Count = 0 Then
    DefaultTemplate = GetDefaultTemplate
    If DefaultTemplate <> "" Then
      Templates.Add DefaultTemplate
    End If
  End If
  Set GetDrawingTemplates = Templates

End Function

Function GetDefaultTemplate() As String

  GetDefaultTemplate = swApp.GetDocumentTemplate(swDocDRAWING, "", swDwgPapersUserDefined, 0, 0)

End Function

Function GetTemplateLocations() As Variant
  
  Dim Locations As String

  Locations = swApp.GetUserPreferenceStringValue(swFileLocationsDocumentTemplates)
  GetTemplateLocations = Split(Locations, ";")

End Function

Sub SearchOpenedDrawing(FullDrawingName As String, ByRef OpenedWindow As ModelWindow)
  
  Dim i As Variant
  Dim AModelWindow As ModelWindow
  Dim DocName As String
  Dim DrawingName As String
  
  Set OpenedWindow = Nothing
  DrawingName = gFSO.GetFileName(FullDrawingName)
  For Each i In swApp.Frame.ModelWindows
    Set AModelWindow = i
    DocName = gFSO.GetFileName(AModelWindow.ModelDoc.GetPathName)
    If StrComp(DrawingName, DocName, vbTextCompare) = 0 Then
      Set OpenedWindow = AModelWindow
      Exit For
    End If
  Next

End Sub

Function CreateDrawingName(ModelFullName As String) As String

  Dim Path As String
  Dim BaseName As String

  Path = gFSO.GetParentFolderName(ModelFullName)
  BaseName = gFSO.GetBaseName(ModelFullName)
  CreateDrawingName = gFSO.BuildPath(Path, BaseName + ".SLDDRW")

End Function

Sub OpenDrawing(DrawingName As String)

  Dim Errors As swFileLoadError_e
  Dim Warnings As swFileLoadWarning_e

  swApp.OpenDoc6 DrawingName, swDocDRAWING, swOpenDocOptions_Silent, "", Errors, Warnings
  
End Sub

Sub GetDocNameAndSaveIfNecessary( _
  Doc As ModelDoc2, ByRef IsNewDoc As Boolean, ByRef Name As String)

  Name = Doc.GetPathName
  IsNewDoc = (Name = "")
  If IsNewDoc Then
    Doc.Extension.RunCommand swCommands_SaveAs, Empty
    Name = Doc.GetPathName
  End If

End Sub
