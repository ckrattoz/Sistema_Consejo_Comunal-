VERSION 5.00
Begin VB.Form cResidencia 
   Caption         =   "Constancia de Residencia"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnBack 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   11
      Top             =   6240
      Width           =   1935
   End
   Begin VB.OptionButton Feme 
      Caption         =   "Femenino"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   5520
      Width           =   1695
   End
   Begin VB.OptionButton Masc 
      Caption         =   "Masculino"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2520
      MaskColor       =   &H80000005&
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtAnosViviendo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtDireccion 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox txtProcedencia 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox txtNoCedula 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtEdad 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtNacionalidad 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton btnGenerar 
      Caption         =   "Generar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   10
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Género"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   960
      TabIndex        =   18
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Años Viviendo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección Completa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Procedencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula Identidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Edad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nacionalidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Completo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   7185
      Left            =   0
      Picture         =   "cResidencia.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6825
   End
End
Attribute VB_Name = "cResidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then
SoloNumeros = 0
Else
SoloNumeros = KeyAscii
End If
' teclas especiales permitidas
If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

Private Sub btnBack_Click()
txtNombre.Text = ""
txtNacionalidad.Text = ""
txtEdad.Text = ""
txtNoCedula.Text = ""
txtProcedencia.Text = ""
txtDireccion.Text = ""
txtAnosViviendo.Text = ""
Masc.Value = False
Feme.Value = False

Set objWord = Nothing
Unload Me
Load Principal
Principal.Show

End Sub

Private Sub btnGenerar_Click()
Dim objWord As Word.Application
If (txtNombre.Text = "") Or (Masc.Value = False And Feme.Value = False) Then
    MsgBox ("Hay Campos Vacios, Por favor comprobar"), vbExclamation, ("Cuidado")
  Else
    Set objWord = New Word.Application
    objWord.Visible = True
    If Masc.Value = True Then
        objWord.Documents.Open App.Path & "\templates\CONSTANCIA_RESIDENCIA_H.doc"
    Else
        objWord.Documents.Open App.Path & "\templates\CONSTANCIA_RESIDENCIA_M.doc"
    End If
    objWord.Documents(1).Bookmarks("nombre").Range = UCase(txtNombre.Text)
    objWord.Documents(1).Bookmarks("nacionalidad").Range = UCase(txtNacionalidad.Text)
    objWord.Documents(1).Bookmarks("edad").Range = txtEdad.Text
    objWord.Documents(1).Bookmarks("cedula").Range = txtNoCedula.Text
    objWord.Documents(1).Bookmarks("procedente").Range = UCase(txtProcedencia.Text)
    objWord.Documents(1).Bookmarks("direccion").Range = UCase(txtDireccion.Text)
    objWord.Documents(1).Bookmarks("hace").Range = txtAnosViviendo.Text
    
End If
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub
Private Sub txtNoCedula_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub
Private Sub txtAnosViviendo_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub
