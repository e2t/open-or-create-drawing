VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseTemplateDialog 
   Caption         =   "Выберите шаблон"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850.001
   OleObjectBlob   =   "ChooseTemplateDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseTemplateDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnCancel_Click()

  ExitApp

End Sub

Private Sub BtnRun_Click()

  Run

End Sub

Private Sub ListBoxNames_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

  If Me.ListBoxNames.ListCount > 0 Then
    Run
  End If

End Sub

Private Sub UserForm_Initialize()

  SearchFormInit

End Sub
