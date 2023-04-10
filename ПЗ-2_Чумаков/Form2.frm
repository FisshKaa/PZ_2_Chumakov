VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   OleObjectBlob   =   "Form2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim a As Integer, b As Integer, s As Integer
a = Val(TextBox1)
b = Val(TextBox2)
x = a + b
TextBox3 = Str(s)
End Sub

Private Sub CommandButton2_Click()
UseForm1.Hide
End Sub

Private Sub CommandButton3_Click()
TextBox1 = ""
TextBox2 = ""
TextBox3 = ""
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub
