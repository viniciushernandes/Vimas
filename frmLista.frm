VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLista 
   BorderStyle     =   3  'Diálogo Fixo
   Caption         =   "Vima's Soft - Lista de arquivos"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3855
   Icon            =   "frmLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Padrão Windows
   Begin MSComDlg.CommonDialog CMD 
      Left            =   2400
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configurações de cores"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "Cor do Fundo"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cor da Fonte"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox txtBusca 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Busca"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fso As New FileSystemObject


Private Sub Command1_Click()
    CMD.ShowColor
    CorFonte = CMD.Color
    
    ArqRelatório = App.Path & "\Config.txt"
    Open ArqRelatório For Output As #1
    Print #1, Tab(1); "[COR]=" & CorFonte
    Print #1, Tab(1); "[FUNDO]=" & CorFundo
    Close #1
End Sub

Private Sub Command2_Click()
    CMD.ShowColor
    CorFundo = CMD.Color
    
    ArqRelatório = App.Path & "\Config.txt"
    Open ArqRelatório For Output As #1
    Print #1, Tab(1); "[COR]=" & CorFonte
    Print #1, Tab(1); "[FUNDO]=" & CorFundo
    Close #1
End Sub

Private Sub Form_Load()
    Dim FArq As String
    FArq = Procura_Arquivo("c:\Windows\System", "showtext.txt")
    If FArq = "" Then
        NChave = ""
        frmRegistro.Show vbModal
        If NChave <> "vimassoftshowtext" Then
            MsgBox "Chave inválida.", vbCritical
            Unload Me
            Exit Sub
        End If
        ArqRelatório = "c:\Windows\System\showtext.txt"
        Open ArqRelatório For Output As #1
        Print #1, Tab(1); "OK"
        Close #1
    End If
        
    frmLista.Top = 0
    frmLista.Left = 0
    X = 0
    ArqRelatório = App.Path & "\Config.txt"
    Open ArqRelatório For Input As #2
    While Not EOF(2)
        X = X + 1
        Line Input #2, Registro
        If X = 1 Then
            If Mid(Registro, 1, 5) = "[COR]" Then
                CorFonte = Mid(Registro, 7)
            Else
                CorFonte = &H80000005
            End If
        ElseIf X = 2 Then
            If Mid(Registro, 1, 7) = "[FUNDO]" Then
                CorFundo = Mid(Registro, 9)
            Else
                CorFundo = &H80000005
            End If
        End If
    Wend
    Close #2
    AchaArquivos fso.GetFolder(App.Path & "\Arquivos")
End Sub

Sub AchaArquivos(f As Folder)
    Dim a As File
    List1.Clear
    For Each a In f.Files
        If UCase$(Trim(Mid(a.Name, 1, Len(txtBusca.Text)))) = UCase$(Trim(txtBusca.Text)) Then
            List1.AddItem Mid(a.Name, 1, InStr(1, a.Name, ".") - 1)
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Unload frmPrévia
    Unload frmPainel
End Sub

Private Sub List1_Click()
    On Error Resume Next
    Unload frmPrévia
    Arquivo = List1.Text
    frmPrévia.Show
End Sub

Private Sub List1_DblClick()
    On Error Resume Next
    Unload frmPrévia
    Arquivo = List1.Text
    frmPainel.Show
End Sub

Private Sub txtBusca_Change()
    AchaArquivos fso.GetFolder(App.Path & "\Arquivos")
End Sub
