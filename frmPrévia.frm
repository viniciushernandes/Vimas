VERSION 5.00
Begin VB.Form frmPrévia 
   BorderStyle     =   0  'Nenhum
   Caption         =   "Vima's Soft - Prévia do arquivo"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   Icon            =   "frmPrévia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Padrão Windows
   Begin VB.TextBox txtPrévia 
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
Attribute VB_Name = "frmPrévia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmPrévia.Top = 0
    frmPrévia.Left = 3950
    txtPrévia.ForeColor = CorFonte
    txtPrévia.BackColor = CorFundo
    X = 0
    txtPrévia.Text = ""
    ArqRelatório = App.Path & "\Arquivos\" & Arquivo & ".txt"
    Open ArqRelatório For Input As #2
    While Not EOF(2)
        X = X + 1
        Line Input #2, Registro
        If X = 1 Then
            If Mid(Registro, 1, 7) = "[FONTE]" Then
                'NFonte = Mid(Registro, 9)
                'txtPrévia.FontName = NFonte
            Else
                txtPrévia.Text = txtPrévia.Text & Registro & vbCrLf
            End If
        ElseIf X = 2 Then
            If Mid(Registro, 1, 9) = "[TAMANHO]" Then
                'TFonte = Mid(Registro, 11)
                'txtPrévia.FontSize = TFonte
            Else
                txtPrévia.Text = txtPrévia.Text & Registro & vbCrLf
            End If
        Else
            txtPrévia.Text = txtPrévia.Text & Registro & vbCrLf
        End If
    Wend
    Close #2
End Sub
