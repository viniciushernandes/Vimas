VERSION 5.00
Begin VB.Form frmPr�via 
   BorderStyle     =   0  'Nenhum
   Caption         =   "Vima's Soft - Pr�via do arquivo"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   Icon            =   "frmPr�via.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Padr�o Windows
   Begin VB.TextBox txtPr�via 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmPr�via"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmPr�via.Top = 0
    frmPr�via.Left = 3950
    txtPr�via.ForeColor = CorFonte
    txtPr�via.BackColor = CorFundo
    X = 0
    txtPr�via.Text = ""
    ArqRelat�rio = App.Path & "\Arquivos\" & Arquivo & ".txt"
    Open ArqRelat�rio For Input As #2
    While Not EOF(2)
        X = X + 1
        Line Input #2, Registro
        If X = 1 Then
            If Mid(Registro, 1, 7) = "[FONTE]" Then
                'NFonte = Mid(Registro, 9)
                'txtPr�via.FontName = NFonte
            Else
                txtPr�via.Text = txtPr�via.Text & Registro & vbCrLf
            End If
        ElseIf X = 2 Then
            If Mid(Registro, 1, 9) = "[TAMANHO]" Then
                'TFonte = Mid(Registro, 11)
                'txtPr�via.FontSize = TFonte
            Else
                txtPr�via.Text = txtPr�via.Text & Registro & vbCrLf
            End If
        Else
            txtPr�via.Text = txtPr�via.Text & Registro & vbCrLf
        End If
    Wend
    Close #2
End Sub
