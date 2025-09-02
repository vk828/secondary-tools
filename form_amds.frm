VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_amds 
   Caption         =   "Please Confirm"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   OleObjectBlob   =   "form_amds.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_amds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserResponse As String

Private Sub cmdUpdate_Click()
    UserResponse = "update"
    Me.Hide
End Sub

Private Sub cmdKeep_Click()
    UserResponse = "keep"
    Me.Hide
End Sub

Private Sub cmdExit_Click()
    UserResponse = ""
    Me.Hide
End Sub
