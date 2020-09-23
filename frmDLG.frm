VERSION 5.00
Begin VB.Form frmDLG 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8430
End
Attribute VB_Name = "frmDLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub File1_DblClick()

Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
FileName = File1.Path & "\" & File1.FileName
MsgBox FileName
End Sub
