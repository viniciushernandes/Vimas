VERSION 5.00
Begin VB.Form frmPainel 
   BorderStyle     =   0  'Nenhum
   Caption         =   "Vima's Soft - Painel de visualização"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   Icon            =   "frmPainel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Padrão Windows
   WindowState     =   2  'Maximizado
   Begin VB.TextBox txtPainel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmPainel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo Erro
    txtPainel.ForeColor = CorFonte
    txtPainel.BackColor = CorFundo
    txtPainel.Text = ""
    X = 0
    ArqRelatório = App.Path & "\Arquivos\" & Arquivo & ".txt"
    Open ArqRelatório For Input As #2
    While Not EOF(2)
        X = X + 1
        Line Input #2, Registro
        If X = 1 Then
            If Mid(Registro, 1, 7) = "[FONTE]" Then
                NFonte = Mid(Registro, 9)
                txtPainel.FontName = NFonte
            Else
                txtPainel.Text = txtPainel.Text & Registro & vbCrLf
            End If
        ElseIf X = 2 Then
            If Mid(Registro, 1, 9) = "[TAMANHO]" Then
                TFonte = Mid(Registro, 11)
                txtPainel.FontSize = TFonte
            Else
                txtPainel.Text = txtPainel.Text & Registro & vbCrLf
            End If
        Else
            txtPainel.Text = txtPainel.Text & Registro & vbCrLf
        End If
    Wend
    Close #2
    Exit Sub
Erro:
    MsgBox "Ocorreu um erro." & vbCrLf & "Verifique o arquivo que está tentando abrir.", vbCritical
    Close #2
    Unload Me
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

