VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ChatGPT"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8730.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
End Sub

Private Sub btnCancel_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub




Private Sub btnCopy_Click()
    
    ClipBoard_SetData lblRephrasedText.text
    ' Update the button caption
    btnCopy.Caption = "Copied!"
    btnCopy.BackColor = &H6B882C
        ' Pause for 2 seconds using a loop
    Dim endTime As Double
    endTime = Now + TimeValue("00:00:04")
    Do While Now < endTime
        DoEvents
    Loop
    btnCopy.Caption = "Copy"
    btnCopy.BackColor = &H97753F
End Sub

Private Sub btnOK_Click()
        Me.Tag = "Insert"
        Me.Hide

End Sub

Private Sub lblRephrasedText_Click()

End Sub

Private Sub UserForm_Click()

End Sub
