VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeclaracaoAnualDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************
'Sistemap Desenvolvimento de Sistemas
'Aplicação: Declaração Anual de Débitos Compesa A4
'Autor: Ericsson Alves
'email: aziel@sistemap.com.br / erinhoeam@hotmail.com
'13 Abril de 2010
'********************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Declaração Quitação Anual de Débitos"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String
Private Type TipoReg  'Dados do Beneficiario e Fatura
        ANO As String
        NOME As String
        MATRICULA As String
        ENDERECO As String
        SEQUENCIAL As String
        INSCRICAO As String
        DESTINATARIO As String
        ENDERECO_DESTINATARIO As String
        BAIRRO As String
        CIDADE As String
        UF As String
        CEP As String
        SEQ As String
        GRUPO As String
        FIRMA As String
        FATURA(12) As String
        SITUACAO(12) As String
        DATA_PAGAMENTO(12) As String
        VALOR(12) As String
        TOTAL_VALOR As String
        ANO_REFERENCIA As String
End Type
'Variáveis gerais para os tipos
Private Reg As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Integer
        Dim posY As Long
        Dim posX As Long
        
        wRec = wRec + 1
        posX = 200
        posY = 440
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 20 scalefont setfont"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Reg.ANO & ") show"
        posY = posY + 270
        Print #outFN, "/Helvetica-Condensed-AP findfont 11 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 180
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        posY = posY + 180
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 120
        
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 12 scalefont setfont"
        posY = posY + 330
        For i = 1 To 12
            Print #outFN, posX - 150 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FATURA(i))) & Reg.FATURA(i) & ") show"
            Print #outFN, posX + 410 & " pontos -" & posY & " pontos moveto (" & Reg.SITUACAO(i) & ") show"
            Print #outFN, posX + 1150 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_PAGAMENTO(i) & ") show"
            Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR(i))) & Reg.VALOR(i) & ") show"

            posY = posY + 40
        Next
        posY = posY + 180
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 14 scalefont setfont"
        Print #outFN, posX + 1560 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_VALOR)) & Reg.TOTAL_VALOR & ") show"
        posY = posY + 230
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 12 scalefont setfont"
        Print #outFN, posX + 1960 & " pontos -" & posY & " pontos moveto (" & Reg.ANO_REFERENCIA & ".) show"
        Print #outFN, "showpage"
        'Fim da Parte Interna
        '**********Inicio Parte Externa************************************************
        posX = 810
        posY = 590
        Print #outFN, "save EPSForm2 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.FIRMA & " - " & Reg.GRUPO & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        posY = posY + 38
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 38
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO_DESTINATARIO & ") show"
        posY = posY + 38
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & "      " & Reg.CIDADE & ") show"
        Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg.UF & ") show"
        posY = posY + 38
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.CEP & ") show"
        Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg.SEQ & ") show"
        posY = posY + 120
        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & " Arq.: " & wArq & " SCDS) show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

        'If wRec = 10 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String
        Dim i1 As Integer
        
        wRec = 0
        totRegs = 0

        'Gravar Cabeçalho
        GravaHeader

        wArq = frmPrincipal.LstSpool.List(frmPrincipal.LstSpool.ListIndex)

        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha
           
           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")
           
           Reg.ANO = Trim(Mid(linha, 1, 4))
           Reg.NOME = Trim(Mid(linha, 5, 50))
           Reg.ENDERECO = Trim(Mid(linha, 55, 120))
           Reg.MATRICULA = Trim(Mid(linha, 175, 9))
           Reg.SEQUENCIAL = Trim(Mid(linha, 184, 50))
           Reg.INSCRICAO = Trim(Mid(linha, 234, 20))
           Reg.FIRMA = Trim(Mid(linha, 254, 10))
           Reg.GRUPO = Trim(Mid(linha, 264, 2))
           Reg.DESTINATARIO = Trim(Mid(linha, 266, 50))
           Reg.ENDERECO_DESTINATARIO = Trim(Mid(linha, 316, 70))
           Reg.BAIRRO = Trim(Mid(linha, 386, 30))
           Reg.CIDADE = Trim(Mid(linha, 416, 30))
           Reg.UF = Trim(Mid(linha, 446, 2))
           Reg.CEP = Trim(Mid(linha, 448, 10))
           Dim winit As Integer
           
           winit = 458
           
           For i1 = 1 To 12
               Reg.FATURA(i1) = Trim(Mid(linha, winit, 7))
               winit = winit + 7
               Reg.SITUACAO(i1) = Trim(Mid(linha, winit, 30))
               winit = winit + 30
               Reg.DATA_PAGAMENTO(i1) = Trim(Mid(linha, winit, 10))
               winit = winit + 10
               Reg.VALOR(i1) = Trim(Mid(linha, winit, 14))
               winit = winit + 14
           Next
           Reg.TOTAL_VALOR = Trim(Mid(linha, 1190, 14))
           Reg.ANO_REFERENCIA = Trim(Mid(linha, 1204, 4))

           MontaExtrato
           LimpaReg
           
           totRegs = totRegs + 1
           frmPrincipal.txQtdReg.Text = totRegs

         '  If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
         '     frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        Close #24
        Close #inFN
        Screen.MousePointer = vbDefault

        frmResultados.txtTotCli_STL.Text = totRegs
        frmResultados.txtTotReg_STL.Text = wRec
        frmResultados.txArquivo.Text = wArq

        frmResultados.txtAplicacao = AppDesc

        Close #outFN
        Close #29

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
       Open iFile For Random As #inFN Len = 1

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
        Open App.Path & "\DeclaracaoQuitacaoAnualDebito.ps" For Input As #200
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
        'Segunda Imagem
        Print #outFN, "/ImageData2"
        Print #outFN, "currentfile"
        Print #outFN, "<< /Filter /SubFileDecode"
        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
        Print #outFN, ">> /ReusableStreamDecode filter"
        Open App.Path & "\DeclaracaoQuitacaoAnualDebito_Ext.ps" For Input As #200
        While Not EOF(200)
            Line Input #200, wCabeca
            Print #outFN, wCabeca
        Wend
        Close #200
        Print #outFN, "*EOD*"
        Print #outFN, "def"
        Print #outFN, " "
        Print #outFN, "/EPSForm2"
        Print #outFN, "<< /FormType 1"
        Print #outFN, "   /BBox [0 0 595 900]"
        Print #outFN, "   /Matrix [ 1 0 0 1 0 0]"
        Print #outFN, "   /PaintProc"
        Print #outFN, "   { pop"
        Print #outFN, "       /ostate save def"
        Print #outFN, "         /showpage {} def"
        Print #outFN, "         /setpagedevice /pop load def"
        Print #outFN, "         ImageData2 0 setfileposition ImageData2 cvx exec"
        Print #outFN, "       ostate restore"
        Print #outFN, "   } bind"
        Print #outFN, ">> def"
        Print #outFN, "%% Fim da Segunda Imagem"
        Print #outFN, " "
        '-------------------------------------------------------------------
        Print #outFN, "%%============================"
        Print #outFN, "%% Inicio dos Dados Variaveis"
        Print #outFN, "%%============================"
        Print #outFN, " "
        Print #outFN, "%Aplicação: Declaração Anual de Quitação de Débito"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer

        Reg.ANO = ""
        Reg.NOME = ""
        Reg.MATRICULA = ""
        Reg.ENDERECO = ""
        Reg.SEQUENCIAL = ""
        Reg.DESTINATARIO = ""
        Reg.ENDERECO_DESTINATARIO = ""
        Reg.BAIRRO = ""
        Reg.CIDADE = ""
        Reg.UF = ""
        Reg.CEP = ""
        Reg.SEQ = ""
        Reg.GRUPO = ""
        Reg.FIRMA = ""
        Reg.INSCRICAO = ""
        For i = 1 To 12
            Reg.FATURA(i) = ""
            Reg.SITUACAO(i) = ""
            Reg.DATA_PAGAMENTO(i) = ""
            Reg.VALOR(i) = ""
        Next
        Reg.TOTAL_VALOR = ""
        Reg.ANO_REFERENCIA = ""
End Sub
