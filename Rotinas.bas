Attribute VB_Name = "Rotinas"
Option Explicit

Public Arquivo As String
Public ArqRelatório As String
Public Registro As String
Public X As Integer
Public NFonte As String
Public TFonte As Integer
Public CorFonte As String
Public CorFundo As String
Public NChave As String

Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, _
ByVal lpInputName As String, ByVal lpOutputName As String) As Long

Public Const MAX_PATH = 260


Public Function Procura_Arquivo(Caminho As String, NomeArquivo As String) As String

Dim lNullPos As Long
Dim lResultado As Long
Dim sBuffer As String

On Error GoTo Procura_Arquivo_Error

'Aloca espaco para a string sBuffer
sBuffer = Space(MAX_PATH * 2)
'inicia busca do arquivo
lResultado = SearchTreeForFile(Caminho, NomeArquivo, sBuffer)

' Se houver um caracter Nulo , remove
If lResultado Then
   lNullPos = InStr(sBuffer, vbNullChar)
    If Not lNullPos Then
       sBuffer = Left(sBuffer, lNullPos - 1)
    End If
   'Retorna o nome do arquivo encontrado
    Procura_Arquivo = sBuffer
Else
    'nao achou nada
    Procura_Arquivo = vbNullString
End If

Exit Function
Procura_Arquivo_Error:
    Procura_Arquivo = vbNullString
End Function



