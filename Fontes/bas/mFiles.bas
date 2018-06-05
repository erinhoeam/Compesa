Attribute VB_Name = "mFiles"
'*****************************************************
'Desenvolvido por Aziel R. Pereira Jr
'                 Analista de Sistemas
'Data: Agosto/2002
'Contatos - Email: aziel@hotmail.com
'           Fone: (81) 9979.1972
'Aplicação: Módulo de Funções
'*****************************************************
Option Explicit

'*****************************************************
'Constantes (flags) para criação e deleção de arquivos
'*****************************************************
Public Const CRIAR = 0
Public Const DELETAR = 1
Public Const EXISTEPASTA = 2
Public Const MOVER = 3
Public Const CRIARPASTA = 4
Public Const EXISTEARQ = 5

'Declarações de Funções API para ler e escrever em arquivos INI
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

'Declarações de Funções API para verificar se o arquivo já está aberto
Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

'Função para criar e deletar arquivos
'Retornará 0 se passou um parâmetro certo, 1 se for passado algo errado
Public Function Manip_Arq(flag As Byte, fName As String) As Boolean
   Dim fso As Object
   Dim op As Object
    
   On Error GoTo Erro
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   If flag = CRIAR Then
      Set op = fso.CreateTextFile(fName, True)
      op.Close
   ElseIf flag = DELETAR Then
      Set op = fso.GetFile(fName)
      op.Delete
   ElseIf flag = EXISTEPASTA Then
      Set op = fso.GetFolder(fName)
   ElseIf flag = CRIARPASTA Then
      If Not fso.FolderExists(fName) Then _
         Set op = fso.CreateFolder(fName)
   ElseIf flag = EXISTEARQ Then
      Set op = fso.GetFile(fName)
   End If
   Manip_Arq = True
   Set op = Nothing
   Set fso = Nothing
   Exit Function
Erro:
   Manip_Arq = False
   Exit Function
End Function

Public Function readINI(AppName As String, KeyName As String, fName As String)
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   readINI = Left(RetStr, GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), fName))
End Function

Public Sub writeINI(AppName As String, KeyName As String, Value As String, fName As String)
   Dim result As Long
   result = WritePrivateProfileString(AppName, KeyName, Value, fName)
   If result <> 1 Then MsgBox "Erro escrevendo no arquivo de configuração."
End Sub

Public Function Aberto(Filename As String) As Boolean
   Dim hFile As Long
   Dim lastErr As Long

   ' Inicializando o handle do arquivo e a variável de erro.
   hFile = -1
   lastErr = 0

   ' Abrir para leitura e compartilhamento exclusivo.
   hFile = lopen(Filename, &H10)

   'Se não for possível abrir o arquivo, pegar o último erro.
   If hFile = -1 Then
      lastErr = Err.LastDllError
   'Lembrar de fechar o arquivo se tudo for ok.
   Else
      lclose (hFile)
   End If

   'Verificar erro de compartilhamento.
   If (hFile = -1) And (lastErr = 32) Then
      Aberto = True
   Else
      Aberto = False
   End If
   
End Function
