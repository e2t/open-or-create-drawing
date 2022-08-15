VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectExistDrawingDialog 
   Caption         =   "Choose a drawing"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850.001
   OleObjectBlob   =   "SelectExistDrawingDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectExistDrawingDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBtn_Click()

  ExitApp

End Sub

Private Sub DrawingListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

  RunChooseExistDrawing

End Sub

Private Sub OkBtn_Click()

  RunChooseExistDrawing

End Sub

Private Sub UserForm_Initialize()

  ChooseDrawingFormInit

End Sub
