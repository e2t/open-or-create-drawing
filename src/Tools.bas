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

  Dim I As Variant
  Dim J As Variant
  Dim aFolder As Folder
  Dim AFile As File
  Dim Extension As String
  Dim Templates As Collection
  Dim DefaultTemplate As String
  
  Set Templates = New Collection
  For Each I In GetTemplateLocations
    Set aFolder = gFSO.GetFolder(I)
    For Each J In aFolder.Files
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
  
  Dim I As Variant
  Dim AModelWindow As ModelWindow
  Dim DocName As String
  Dim DrawingName As String
  
  Set OpenedWindow = Nothing
  DrawingName = gFSO.GetFileName(FullDrawingName)
  For Each I In swApp.Frame.ModelWindows
    Set AModelWindow = I
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

' Без точек "." в наименовании
Sub SplitNameAndDesignation(Src As String, ByRef Designation As String, _
                            ByRef Name As String)

  Const Flat As String = "SM-FLAT-PATTERN"
  Const gCodeRegexPattern = "СБ|МЧ|УЧ|ВО|РСБ|AD|ID|\.AD|\.ID"
  Dim RegexAsm As RegExp
  Dim RegexPrt As RegExp
  Dim Matches As Object
  Dim Z As Variant
  
  Designation = Src
  Name = Src
  'Code = ""
  
  Set RegexAsm = New RegExp
  RegexAsm.Pattern = "(.*\..*[0-9] *)(" + gCodeRegexPattern + ") ([^.]+)"
  RegexAsm.IgnoreCase = True
  RegexAsm.Global = True
  
  Set RegexPrt = New RegExp
  RegexPrt.Pattern = "(.*\.[^ ]+) ([^.]+)"
  RegexPrt.IgnoreCase = True
  RegexPrt.Global = True
  
  If RegexAsm.Test(Src) Then
    Set Matches = RegexAsm.Execute(Src)
    Designation = Trim(Matches(0).SubMatches(0))
    'Code = Matches(0).SubMatches(1)
    Name = Trim(Matches(0).SubMatches(2))
  ElseIf RegexPrt.Test(Src) Then
    Set Matches = RegexPrt.Execute(Src)
    Designation = Trim(Matches(0).SubMatches(0))
    Name = Trim(Matches(0).SubMatches(1))
  End If
   
End Sub

Function IsBaseConf(Conf As String) As Boolean

  Select Case Conf
    Case "00", "По умолчанию", "Default"
      IsBaseConf = True
    Case Else
      IsBaseConf = False
  End Select

End Function

Function SearchDrawings(Designation As String, FolderPath As String) As Dictionary

  Dim aFolder As Folder
  Dim I As Variant
  Dim F As File
  Dim RegexDrawing As RegExp
  Dim Result As Dictionary
  
  Set Result = New Dictionary
  
  Set RegexDrawing = New RegExp
  RegexDrawing.Pattern = "( *" + RegEscape(Designation) + " +)(.*)\.SLDDRW"
  RegexDrawing.IgnoreCase = True
  RegexDrawing.Global = True

  Set aFolder = gFSO.GetFolder(FolderPath)
  For Each I In aFolder.Files
    Set F = I
    If RegexDrawing.Test(F.ShortName) Then
      Result.Add gFSO.GetBaseName(F.ShortName), F.Path
    End If
  Next
  Set SearchDrawings = Result

End Function

'See: https://docs.microsoft.com/en-us/dotnet/api/System.Text.RegularExpressions.Regex.Escape
Function RegEscape(ByVal Line As String) As String

  Line = Replace(Line, "\", "\\")  'MUST be first!
  Line = Replace(Line, ".", "\.")
  Line = Replace(Line, "[", "\[")
  'line = Replace(line, "]", "\]")
  Line = Replace(Line, "|", "\|")
  Line = Replace(Line, "^", "\^")
  Line = Replace(Line, "$", "\$")
  Line = Replace(Line, "?", "\?")
  Line = Replace(Line, "+", "\+")
  Line = Replace(Line, "*", "\*")
  Line = Replace(Line, "{", "\{")
  'line = Replace(line, "}", "\}")
  Line = Replace(Line, "(", "\(")
  Line = Replace(Line, ")", "\)")
  Line = Replace(Line, "#", "\#")
  'and white space??
  RegEscape = Line
    
End Function
