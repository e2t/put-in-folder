VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Переместить в папку"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6090
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonClose_Click()

  ExitApp

End Sub

Private Sub ButtonRun_Click()

  RunIfSelected

End Sub

Private Sub ButtonSettings_Click()

  OpenSettingsFile
  ExitApp

End Sub

Private Sub ListBoxName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

  RunIfSelected

End Sub

Private Sub UserForm_Initialize()

  MainFormInit

End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

  If KeyAscii = 27 Then
    ExitApp
  End If

End Sub
