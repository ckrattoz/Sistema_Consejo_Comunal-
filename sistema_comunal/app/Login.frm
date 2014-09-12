VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Acceso al Sistema Consejo Comunal"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Caption         =   "Acceder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      MaskColor       =   &H00000080&
      TabIndex        =   4
      Top             =   6240
      Width           =   3015
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox txtuser 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SISTEMA CONSEJO COMUNAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   9135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio Sesión"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Index           =   10
      Left            =   3600
      TabIndex        =   5
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   7215
      Index           =   1
      Left            =   0
      Picture         =   "Login.frx":1085C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9465
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtuser.Text = "" Then
MsgBox "El campo de usuario no puede estar vacio.", vbInformation
txtuser.SetFocus
Exit Sub
ElseIf txtpassword.Text = "" Then
MsgBox "El campo de contrseña no puede estar vacio"
txtpassword.SetFocus
Exit Sub
Else
Call login
End If
End Sub

Private Sub login()
Module1.getConnected
Dim rs As New ADODB.Recordset
rs.Open "Select * From tblusers Where username = '" & txtuser.Text & "'", cnn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
MsgBox "El nombre de usuario no es valido. Por favor intente de nuevo.", vbInformation, "Login"
txtuser.SetFocus
Exit Sub
Else
If txtpassword.Text = rs!password Then
Unload Me
Load Principal
Principal.Show
Exit Sub
Else
MsgBox "Contraseña no valida. Por favor intente de nuevo.", vbInformation, "Login"

txtpassword.SetFocus
Exit Sub
End If
End If
Set rs = Nothing
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub
