VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   OleObjectBlob   =   "for_Debug.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub from_p_Change()

End Sub

Private Sub in_Q_Change()

End Sub


Private Sub Translate_box_Change()

End Sub

Private Sub Result_Change()

End Sub

Private Sub cmd_Clear_Click()
txt_A = ""
txt_B = ""
txt_P = ""
txt_Q = ""
End Sub

Private Sub cmd_Exitr_Click()
UserForm1.Hide
End Sub

Private Sub cmd_OK_Click()
Dim num_P, num_Q, num_10, num_S
num_P = Val(txt_P)
num_Q = Val(txt_Q)
For i = 1 To Len(txt_A)
num_10 = num_10 + Val(Mid(txt_A, i, 1)) * num_P ^ (Len(txt_A) - i)
Debug.Print ("num_10=" & num_10)
Next i
txt_B = ""
While num_10 <> 0
num_S = num_10 Mod num_Q
Debug.Print ("num_S" & num_S)
txt_B = Mid(Str(num_S), 2, 1) + txt_B
Debug.Print ("txt_B=" & txt_B)
num_10 = num_10 \ num_Q
Wend
End Sub

Private Sub txt_P_Change()

End Sub
