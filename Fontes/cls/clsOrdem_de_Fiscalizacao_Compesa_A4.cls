VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdem_de_Fisc_A4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap Consultoria
'                 Analista de Sistemas
'Data: 20 de Agosto de 2008
'Contatos - Email: erinhoeam@hotmail.com
'                  aziel@sistemap.com.br
'           Fone: (81) 9832.6084
'Aplicação: Ordem de Fiscalização Compesa A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double
Dim wQtdFaturas As Long
Dim wIndiceFaturas As Long

Const AppName = "Compesa", AppDesc = "Ordem de Fiscalização Compesa A4"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        EXECUTOR As String
        ESCRITORIO As String
        INSCRICAO As String
        NOME As String
        MATRICULA As String
        ENDERECO As String
        NR_ORDEM As String
        GRUPO As String
        RES_COM_IND_PUB As String
        TOTAL_ECONOMIA As String
        DATA_EMISSAO As String
        TIPO_CONSUMIDOR As String
        ULTIMA_ALTERACAO As String
        SEQUENCIA As String
        SITUACAO_AGUA As String
        CONSUMO_MEDIO As String
        SITUACAO_ESGOTO As String
        CONSUMO_FIXO As String
        DATA_INTERRUPCAO As String
        SUPESSAO_PARCIAL As String
        SUPESSAO_TOTAL As String
        ORIGEM As String
        OCORRENCIA As String
        FATURA() As String
        VENCIMENTO() As String
        VALOR() As String
        DATA_DEBITO As String
        VALOR_TOTAL As String
End Type
'Variáveis gerais para os tipos
Private Reg As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim posY As Long
        Dim posX As Long

        posX = 100
        posY = 215

        wRec = wRec + 1
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE

       'Executor, Escritorio
        Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg.EXECUTOR & ") show"
        Print #outFN, posX + 1150 & " pontos -" & posY & " pontos moveto (" & Reg.ESCRITORIO & ") show"
        posY = posY + 90
        'Inscrição, Nome
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 610 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        posY = posY + 80
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        'Endereço, Grupo, Numero de Ordem
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg.NR_ORDEM & ") show"
        posY = posY + 85
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 90 & " pontos -" & posY & " pontos moveto (" & Reg.RES_COM_IND_PUB & ") show"
        Print #outFN, posX + 445 & " pontos -" & posY & " pontos moveto (" & Reg.TOTAL_ECONOMIA & ") show"
        Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO & ") show"
        Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg.TIPO_CONSUMIDOR & ") show"
        Print #outFN, posX + 1070 & " pontos -" & posY & " pontos moveto (" & Reg.ULTIMA_ALTERACAO & ") show"
        Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO & ") show"
        Print #outFN, posX + 1480 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIA & ") show"
        posY = posY + 85
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.SITUACAO_AGUA & ") show"
        Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_MEDIO & ") show"
        Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg.SITUACAO_ESGOTO & ") show"
        posY = posY + 90
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_INTERRUPCAO & ") show"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.SUPESSAO_PARCIAL & ") show"
        Print #outFN, posX + 770 & " pontos -" & posY & " pontos moveto (" & Reg.ORIGEM & ") show"
        Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Reg.OCORRENCIA & ") show"
        
        posY = posY + 140
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 8 scalefont setfont"
        For i = 1 To wQtdFaturas
            Print #outFN, posX + 250 & " pontos -" & posY & " pontos moveto (" & Reg.FATURA(i) & ") show"
            Print #outFN, posX + 670 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO(i) & ") show"
            Print #outFN, posX + 1000 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.VALOR(i))) & Reg.VALOR(i) & ") show"
            posY = posY + 30

        Next
       'Total Débito
        posY = 1250
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_DEBITO & ") show"
        Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.VALOR_TOTAL)) & Reg.VALOR_TOTAL & ") show"
        posX = 1800
        posY = 215
       'Executor, Escritorio
        Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
        Print #outFN, posX + 1150 & " pontos -" & posY & " pontos moveto (Escritorio: " & Reg.ESCRITORIO & ") show"
        posY = posY + 90
        'Inscrição, Nome
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 645 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        posY = posY + 80
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        'Endereço, Grupo, Numero de Ordem
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg.NR_ORDEM & ") show"
        posY = posY + 85
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 90 & " pontos -" & posY & " pontos moveto (" & Reg.RES_COM_IND_PUB & ") show"
        'Print #outFN, posX + 445 & " pontos -" & posY & " pontos moveto (" & Reg.TOTAL_ECONOMIA & ") show"
        Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO & ") show"
        Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg.VALOR_TOTAL & ") show"
        Print #outFN, posX + 1070 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_DEBITO & ") show"
        Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO & ") show"
        Print #outFN, posX + 1480 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIA & ") show"
        
        posY = posY + 755
        
        Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
        Print #outFN, posX + 1150 & " pontos -" & posY & " pontos moveto (Escritorio: " & Reg.ESCRITORIO & ") show"
        posY = posY + 90
        'Inscrição, Nome
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 645 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        posY = posY + 80
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        'Endereço, Grupo, Numero de Ordem
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg.NR_ORDEM & ") show"
        posY = posY + 85
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 90 & " pontos -" & posY & " pontos moveto (" & Reg.RES_COM_IND_PUB & ") show"
        'Print #outFN, posX + 445 & " pontos -" & posY & " pontos moveto (" & Reg.TOTAL_ECONOMIA & ") show"
        Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO & ") show"
        Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg.VALOR_TOTAL & ") show"
        Print #outFN, posX + 1070 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_DEBITO & ") show"
        Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO & ") show"
        Print #outFN, posX + 1480 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIA & ") show"
        posY = posY + 990
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        
'            'Escritorio, Sequencia
'            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
'            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESCRITORIO2 & ") show"
'            posY = posY + 530
'            'Inscrição, Nome, Matricula
'            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
'            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
'            Print #outFN, posX + 580 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
'            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
'            Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
'            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
'            posY = posY + 95
'            'Endereço, Grupo, Numero de Ordem
'            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
'            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
'            Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg(i).NR_ORDEM & ") show"
'            posY = posY + 80
'            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
'            'Com, ind, Pub, Tipo Consumidor, Data Emissão
'            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).RES_COM_IND_PUB & ") show"
'            Print #outFN, posX + 440 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_EMISSAO & ") show"
'            Print #outFN, posX + 670 & " pontos -" & posY & " pontos moveto (" & Reg(i).TIPO_CONSUMIDOR & ") show"
'            Print #outFN, posX + 1000 & " pontos -" & posY & " pontos moveto (" & Reg(i).A_E & ") show"
'            Print #outFN, posX + 1170 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
'            Print #outFN, posX + 1400 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQUENCIA & ") show"
'            posY = posY + 180
'            'Sequencia e Recuperação
'            Print #outFN, "/Helvetica-AP findfont 4.5 scalefont setfont"
'            Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (Rec: " & Format(wRec, "000000") & "   Seq: " & Reg(i).SEQUENCIA & ") show"
        
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "
        'If wRec = 10 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wContLinha As Integer, wFlagFaturas As Boolean
        wFlagFaturas = False
        wQtdFaturas = 0
        wContLinha = 0
        totRegs = 0

        'Dividir Registros
        Open pastaProcessar & "temporario.tmp" For Output As #24
        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha

           If Mid(linha, 1, 2) = "14" Then
              Print #24, linha
              Print #24, ""
           Else
              Print #24, linha
           End If

         '  If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
         '     frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        Close #24
        Close #inFN

        totRegs = 0

        'Gravar Cabeçalho
        GravaHeader

        wArq = frmPrincipal.LstSpool.List(frmPrincipal.LstSpool.ListIndex)
        Open pastaProcessar & "temporario.tmp" For Input As #29
        'frmPrincipal.ProgressBarProcess.Max = LOF(29)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001

        Do While Not EOF(29)
           DoEvents
           Line Input #29, linha
           
           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")
           wContLinha = wContLinha + 1

           Select Case wContLinha
                Case 2
                    Reg.EXECUTOR = Trim(Mid(linha, 29, 55))
                    Reg.ESCRITORIO = Trim(Mid(linha, 84, 40))
                Case 3
                    Reg.INSCRICAO = Trim(Mid(linha, 5, 20))
                    Reg.NOME = Trim(Mid(linha, 37, 50))
                Case 4
                    Reg.MATRICULA = Trim(Mid(linha, 71, 10))
                Case 7
                    Reg.ENDERECO = Trim(Mid(linha, 5, 80))
                Case 8
                    Reg.NR_ORDEM = Trim(Mid(linha, 74, 7))
                Case 10
                    Reg.RES_COM_IND_PUB = Trim(Mid(linha, 5, 22))
                    Reg.TOTAL_ECONOMIA = Trim(Mid(linha, 27, 4))
                    Reg.DATA_EMISSAO = Trim(Mid(linha, 34, 10))
                    Reg.TIPO_CONSUMIDOR = Trim(Mid(linha, 48, 16))
                    Reg.ULTIMA_ALTERACAO = Trim(Mid(linha, 64, 10))
                    Reg.GRUPO = Trim(Mid(linha, 80, 4))
                    Reg.SEQUENCIA = Trim(Mid(linha, 86, 7))
                    Reg.SEQUENCIA = Trim(Mid(linha, 86, 7))
                Case 11
                    Reg.SITUACAO_AGUA = Trim(Mid(linha, 5, 15))
                    Reg.CONSUMO_MEDIO = Trim(Mid(linha, 21, 5))
                    Reg.SITUACAO_ESGOTO = Trim(Mid(linha, 53, 20))
                Case 12
                    Reg.DATA_INTERRUPCAO = Trim(Mid(linha, 6, 10))
                    Reg.SUPESSAO_PARCIAL = Trim(Mid(linha, 22, 10))
                    Reg.ORIGEM = Trim(Mid(linha, 45, 15))
                    Reg.OCORRENCIA = Trim(Mid(linha, 60, 50))
            End Select
            If Mid(linha, 1, 2) = "12" Then
               wFlagFaturas = False
            End If
            If Mid(linha, 1, 2) = "02" Then
               Reg.DATA_DEBITO = Trim(Mid(linha, 27, 10))
               Reg.VALOR_TOTAL = Trim(Mid(linha, 66, 14))
            End If
            If Mid(linha, 1, 2) = "01" Then
               wFlagFaturas = True
            End If
            If wFlagFaturas Then
               wQtdFaturas = wQtdFaturas + 1
               ReDim Preserve Reg.FATURA(wQtdFaturas)
               ReDim Preserve Reg.VENCIMENTO(wQtdFaturas)
               ReDim Preserve Reg.VALOR(wQtdFaturas)
               
               Reg.FATURA(wQtdFaturas) = Trim(Mid(linha, 14, 9))
               Reg.VENCIMENTO(wQtdFaturas) = Trim(Mid(linha, 40, 10))
               Reg.VALOR(wQtdFaturas) = Trim(Mid(linha, 65, 14))
            End If


           If Mid(linha, 1, 1) = "" Then
              wContLinha = 0
              MontaExtrato
              totRegs = totRegs + 1
              frmPrincipal.txQtdReg.Text = totRegs
              LimpaReg
              wQtdFaturas = 0
           End If

           'If Seek(29) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(29)
        Loop
        Screen.MousePointer = vbDefault

        frmResultados.txtTotCli_STL.Text = totRegs
        frmResultados.txtTotReg_STL.Text = wRec
        frmResultados.txArquivo.Text = wArq

        frmResultados.txtAplicacao = AppDesc

        Close #outFN
        Close #29

        Kill pastaProcessar & "*.tmp"

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
        Open App.Path & "\Ordem_de_Fiscalizacao_Compesa_A4.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Ordem de Fiscalização Compesa A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Reg.EXECUTOR = ""
        Reg.ESCRITORIO = ""
        Reg.INSCRICAO = ""
        Reg.NOME = ""
        Reg.MATRICULA = ""
        Reg.ENDERECO = ""
        Reg.NR_ORDEM = ""
        Reg.GRUPO = ""
        Reg.RES_COM_IND_PUB = ""
        Reg.TOTAL_ECONOMIA = ""
        Reg.DATA_EMISSAO = ""
        Reg.TIPO_CONSUMIDOR = ""
        Reg.ULTIMA_ALTERACAO = ""
        Reg.SEQUENCIA = ""
        Reg.SITUACAO_AGUA = ""
        Reg.CONSUMO_MEDIO = ""
        Reg.SITUACAO_ESGOTO = ""
        Reg.CONSUMO_FIXO = ""
        Reg.DATA_INTERRUPCAO = ""
        Reg.SUPESSAO_PARCIAL = ""
        Reg.SUPESSAO_TOTAL = ""
        Reg.ORIGEM = ""
        Reg.OCORRENCIA = ""
        ReDim Reg.FATURA(0) As String
        ReDim VENCIMENTO(0) As String
        ReDim VALOR(0) As String
        Reg.DATA_DEBITO = ""
        Reg.VALOR_TOTAL = ""
End Sub
