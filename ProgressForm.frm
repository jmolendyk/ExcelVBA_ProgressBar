VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Progress"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4395
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InterruptButton_Click()
    Application.SendKeys "^{BREAK}"
End Sub

Private Sub UserForm_Initialize()
    ProgressForm.Status.Caption = "0% Complete"
    ProgressForm.Bar.Width = 10
End Sub


Sub start(Optional strCaption As String = "Progress", Optional bInterrupt As Boolean = False)
    InterruptButton.Visible = bInterrupt
    ProgressForm.Caption = strCaption
    ProgressForm.Show (0)
End Sub


Sub update(lngPercentComplete As Long, Optional strStatus As String = "")

    ProgressForm.Bar.Width = 2 * lngPercentComplete
    ProgressForm.Status.Caption = lngPercentComplete & "% Complete"

    SubStatus.Caption = strStatus
    DoEvents
End Sub

Sub done()
  Unload ProgressForm
End Sub

