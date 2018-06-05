VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifa_Social_Compesa_A5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap Consultoria
'                 Analista de Sistemas
'Data: 27 de Novembro de 2008
'Contatos - Email: aziel@sistemap.com.br
'                  erinhoeam@hotmail.com
'           Fone: (81) 9832.6084
'Aplicação: Nova Carta de Tarifa Social A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Nova Carta de Tarifa Social A4/A5"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        INSCRICAO As String
        MATRICULA As String
        NOME As String
        ENDERECO As String
        GRUPO As String
        EMPRESA As String
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Integer
        Dim posY As Long
        Dim posX As Long
        
        wRec = wRec + 1

        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        'Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        
        For i = 1 To 2
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            If Len(Trim(Reg(i).NOME)) > 0 Then totRegs = totRegs + 1
            Select Case i
                Case 1
                    posX = 520
                    posY = 458
                Case 2
                    posX = 520
                    posY = 2208
            End Select
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 545 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            posY = posY + 75
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            Print #outFN, posX + 545 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
            posY = posY + 1200
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 4.5 scalefont setfont"
            Print #outFN, posX + 370 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & " Rec.: " & Format(wRec, "00000") & ") show"
        Next
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

        'If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String
        Dim wIndice As Integer
        wIndice = 0

        wRec = 0
        totRegs = 0

        'Gravar Cabeçalho
        GravaHeader

        wArq = frmPrincipal.LstSpool.List(frmPrincipal.LstSpool.ListIndex)

        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001

        Do While Not EOF(inFN)
           DoEvents
           Line Input #inFN, linha

           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")

           wIndice = wIndice + 1

           Reg(wIndice).INSCRICAO = Trim(Mid(linha, 1, 20))
           Reg(wIndice).MATRICULA = Trim(Mid(linha, 22, 10))
           Reg(wIndice).NOME = Trim(Mid(linha, 32, 40))
           Reg(wIndice).ENDERECO = Trim(Mid(linha, 72, 60))
           Reg(wIndice).GRUPO = Trim(Mid(linha, 132, 2))
           Reg(wIndice).EMPRESA = Trim(Mid(linha, 134, 2))

           If wIndice = 2 Then
              MontaExtrato
              LimpaReg
              wIndice = 0
           End If

           frmPrincipal.txQtdReg.Text = totRegs

         '  If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
         '     frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        If wIndice > 0 Then MontaExtrato
        frmPrincipal.txQtdReg.Text = totRegs
        
        frmResultados.txtTotCli_STL.Text = totRegs
        frmResultados.txtTotReg_STL.Text = wRec
        frmResultados.txArquivo.Text = wArq

        Close #inFN
        Close #outFN
        Screen.MousePointer = vbDefault

        frmPrincipal.MousePointer = vbDefault
        'frmPrincipal.ProgressBarProcess.Value = 0.0001

   Exit Sub
'--------------------------------------------------------------------------------
Erro:
   Close #outFN
   Close #inFN
   
   frmPrincipal.MousePointer = vbDefault
   'frmPrincipal.ProgressBarProcess.Value = 0.0001
   
   MsgBox "Descrição do Erro: " & Err.Description & vbCrLf & _
          "Número do Erro:    " & Err.Number & vbCrLf & _
          "ERRO durante o processamento do arquivo " & procArq & "." & vbCrLf & _
          "Caso o arquivo " & gerArq & vbCrLf & _
          "ainda exista, favor APAGÁ-LO E FAZER TODO O PROCESSO DO INÍCIO.", vbCritical
   
   If Manip_Arq(EXISTEARQ, gerArq) Then
      If Not Manip_Arq(DELETAR, gerArq) Then
         MsgBox "O arquivo " & gerArq & " não pode ser excluído." & _
         vbCrLf & "Favor apagá-lo e reiniciar o processamento.", vbCritical
      End If
   End If
End Sub
Public Sub DoProcess(iFile As String, pProc As String, pGer As String)
       Dim oFile As String
   
       inFN = FreeFile
       Open iFile For Input As #inFN

       pastaProcessar = pProc
       pastaGerar = pGer

       jobName = nomeDoJob(iFile)

       outFN = FreeFile
       oFile = pastaGerar & Left$(jobName, Len(jobName) - 4) & ".ps"

       jobName = nomeDoJob(oFile)

       Open oFile For Output As outFN

       totRegs = 0

       procArq = iFile
       gerArq = oFile

       Processar
End Sub
Private Sub GravaHeader()
        Dim wCabeca As String
        'Cabeçalho
        Open App.Path & "\Header_Compesa_A4.ps" For Input As #200
        While Not EOF(200)
            Line Input #200, wCabeca
            Print #outFN, wCabeca
        Wend
        Close #200
        Print #outFN, " "
        '---------------------------------------------------------------
        'Primeira Imagem
        Print #outFN, "/ImageData1"
        Print #outFN, "currentfile"
        Print #outFN, "<< /Filter /SubFileDecode"
        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
        Print #outFN, ">> /ReusableStreamDecode filter"
        Open App.Path & "\Nova_Carta_Tarifa_Social.ps" For Input As #200
        While Not EOF(200)
            Line Input #200, wCabeca
            Print #outFN, wCabeca
        Wend
        Close #200
        Print #outFN, "*EOD*"
        Print #outFN, "def"
        Print #outFN, " "
        Print #outFN, "/EPSForm1"
        Print #outFN, "<< /FormType 1"
        Print #outFN, "   /BBox [0 0 595 900]"
        Print #outFN, "   /Matrix [ 1 0 0 1 0 0]"
        Print #outFN, "   /PaintProc"
        Print #outFN, "   { pop"
        Print #outFN, "       /ostate save def"
        Print #outFN, "         /showpage {} def"
        Print #outFN, "         /setpagedevice /pop load def"
        Print #outFN, "         ImageData1 0 setfileposition ImageData1 cvx exec"
        Print #outFN, "       ostate restore"
        Print #outFN, "   } bind"
        Print #outFN, ">> def"
        Print #outFN, "%% Fim da Primeira Imagem"
        Print #outFN, " "
        '-------------------------------------------------------------------
        Print #outFN, "%%============================"
        Print #outFN, "%% Inicio dos Dados Variaveis"
        Print #outFN, "%%============================"
        Print #outFN, " "
        Print #outFN, "%Aplicação: Nova Carta de Tarifa Social Compesa A4/A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer

        For i = 1 To 2
            Reg(i).INSCRICAO = ""
            Reg(i).NOME = ""
            Reg(i).ENDERECO = ""
            Reg(i).MATRICULA = ""
            Reg(i).GRUPO = ""
            Reg(i).EMPRESA = ""
        Next
End Sub
