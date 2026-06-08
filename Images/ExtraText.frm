VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExtraText 
   ClientHeight    =   5076
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10716
   OleObjectBlob   =   "ExtraText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExtraText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackButton_Click()
Dim TextBoxCount As Long
TextBoxCount = 3
If TextBox3.Visible = True Then
TextBox1.Visible = False
TextBox2.Visible = True
TextBox3.Visible = False
NextButton.Visible = True
BackButton.Visible = True
Exit Sub
End If
If TextBox2.Visible = True Then
TextBox1.Visible = True
TextBox2.Visible = False
TextBox3.Visible = False
NextButton.Visible = True
BackButton.Visible = False
End If

End Sub


Private Sub CloseText_Click()
Unload Me
End Sub

Private Sub NextButton_Click()
Dim TextBoxCount As Long
TextBoxCount = 3
If TextBox1.Visible = True Then
TextBox1.Visible = False
TextBox2.Visible = True
TextBox3.Visible = False
NextButton.Visible = True
BackButton.Visible = True
Exit Sub
End If
If TextBox2.Visible = True Then
TextBox1.Visible = False
TextBox2.Visible = False
TextBox3.Visible = True
NextButton.Visible = False
BackButton.Visible = True
End If
End Sub
