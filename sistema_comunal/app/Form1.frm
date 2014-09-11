VERSION 5.00
Begin VB.Form Principal 
   Caption         =   "Sistema Consejo Comunal"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   8655
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9945
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub
