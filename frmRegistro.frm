VERSION 5.00
Begin VB.Form frmRegistro 
   BorderStyle     =   3  'Diálogo Fixo
   Caption         =   "Registro"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtChave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Digite a chave de registro:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1860
   End
End
Attribute VB_Name = "frmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    NChave = txtChave.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
