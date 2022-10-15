VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} roomsList_UserForm 
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4140
   OleObjectBlob   =   "roomsList_UserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "roomsList_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub roomsList_ListBox_Click()

End Sub

Private Sub roomsList_ListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Worksheets(BOOKING_WS_NAME).meetingRooms_ComboBox.value = Me.roomsList_ListBox

End Sub

Private Sub UserForm_Click()

End Sub
