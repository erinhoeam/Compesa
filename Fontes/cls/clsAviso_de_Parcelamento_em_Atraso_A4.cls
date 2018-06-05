VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAviso_de_Parc_em_Atraso_A4"
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
'Aplicação: Aviso de Parcelamento em Atraso Compesa A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Aviso de Parcelamento em Atraso Compesa A4"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        GRUPO_EMPREITEIRA As String
        INSCRICAO As String
        ESCRITORIO As String
        NOME As String
        ENDERECO As String
        BAIRRO As String
        CIDADE As String
        UF As String
        CEP As String
        SEQ As String
        MATRICULA As String
        SEQUENCIAL As String
        ORDEM As String
        LINHA_DIGITAVEL As String
        BAR_CODE As String
        DATA_PARCELAMENTO As String
        VALOR_DEBITO_ORIGINAL As String
        VALOR_ENCARGOS As String
        DESCONTO_CONCEDIDO As String
        VALOR_NEGOCIADO As String
        JUROS_PARCELAMENTO As String
        VALOR_ENTRADA As String
        QTD_PARCELAS As String
        VALOR_PARCELAS As String
        FATURA(15) As String
        VENCIMENTO(15) As String
        VALOR_PRESTACAO(15) As String
        VALOR_TOTAL_CONTA(15) As String
        TOTAL As String
        VENCIMENTO_BOLETO As String
        VALOR As String
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
        posX = 340
        posY = 360

        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        'Rota
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        posY = posY + 117
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 280
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 8 scalefont setfont"
        Print #outFN, posX - 155 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_PARCELAMENTO & ") show"
        Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_DEBITO_ORIGINAL)) & Reg.VALOR_DEBITO_ORIGINAL & ") show"
        Print #outFN, posX + 290 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_ENCARGOS)) & Reg.VALOR_ENCARGOS & ") show"
        Print #outFN, posX + 540 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.DESCONTO_CONCEDIDO)) & Reg.DESCONTO_CONCEDIDO & ") show"
        Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_NEGOCIADO)) & Reg.VALOR_NEGOCIADO & ") show"
        Print #outFN, posX + 1030 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.JUROS_PARCELAMENTO)) & Reg.JUROS_PARCELAMENTO & ") show"
        Print #outFN, posX + 1270 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_ENTRADA)) & Reg.VALOR_ENTRADA & ") show"
        Print #outFN, posX + 1400 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.QTD_PARCELAS)) & Reg.QTD_PARCELAS & ") show"
        Print #outFN, posX + 1660 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_PARCELAS)) & Reg.VALOR_PARCELAS & ") show"
        posY = posY + 270
        For i = 1 To 15
            Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FATURA(i))) & Reg.FATURA(i) & ") show"
            Print #outFN, posX + 520 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO(i) & ") show"
            Print #outFN, posX + 855 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_PRESTACAO(i))) & Reg.VALOR_PRESTACAO(i) & ") show"
            Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_TOTAL_CONTA(i))) & Reg.VALOR_TOTAL_CONTA(i) & ") show"
            
            posY = posY + 35
        Next
        posY = 1590
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL)) & Reg.TOTAL & ") show"
        posY = posY + 575
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 1170 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO_BOLETO & ") show"
        posY = posY + 860
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg.LINHA_DIGITAVEL & ") show"
        posY = posY + 180
        Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.BAR_CODE & ") show"
        posY = posY + 95
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.ORDEM & ") show"
        Print #outFN, posX + 1750 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 120
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 790 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO_BOLETO & ") show"
        Print #outFN, posX + 1400 & " pontos -" & posY & " pontos moveto (" & Reg.VALOR & ") show"
        Print #outFN, "showpage"
        'Fim da Parte Interna
        '**********Inicio Parte Externa************************************************
        posX = 810
        posY = 590
        Print #outFN, "save EPSForm2 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO_EMPREITEIRA & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        posY = posY + 38
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.ESCRITORIO & ") show"
        posY = posY + 38
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 38
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        posY = posY + 38
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & "      " & Reg.CIDADE & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.UF & ") show"
        posY = posY + 38
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.CEP & ") show"
        Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg.SEQ & ") show"
        posY = posY + 80
        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

        'If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String
        Dim i1 As Integer
        Dim i2 As Integer
        Dim i3 As Integer
        Dim i4 As Integer
        Dim i5 As Integer
        
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
           
           Reg.GRUPO_EMPREITEIRA = Trim(Mid(linha, 1, 27))
           Reg.INSCRICAO = Trim(Mid(linha, 28, 20))
           Reg.ESCRITORIO = Trim(Mid(linha, 48, 25))
           Reg.NOME = Trim(Mid(linha, 73, 50))
           Reg.ENDERECO = Trim(Mid(linha, 123, 50))
           Reg.BAIRRO = Trim(Mid(linha, 173, 25))
           Reg.CIDADE = Trim(Mid(linha, 198, 24))
           Reg.UF = Trim(Mid(linha, 222, 2))
           Reg.CEP = Trim(Mid(linha, 224, 14))
           Reg.SEQ = Trim(Mid(linha, 238, 13))
           Reg.MATRICULA = Trim(Mid(linha, 251, 9))
           Reg.SEQUENCIAL = Trim(Mid(linha, 310, 10))
           Reg.VENCIMENTO_BOLETO = Trim(Mid(linha, 320, 10))
           Reg.ORDEM = Trim(Mid(linha, 330, 5))
           Reg.LINHA_DIGITAVEL = Trim(Mid(linha, 335, 55))
           Reg.BAR_CODE = Trim(Mid(linha, 390, 112))
           Reg.DATA_PARCELAMENTO = Trim(Mid(linha, 502, 10))
           Reg.VALOR_DEBITO_ORIGINAL = Trim(Mid(linha, 512, 14))
           Reg.VALOR_ENCARGOS = Trim(Mid(linha, 526, 14))
           Reg.DESCONTO_CONCEDIDO = Trim(Mid(linha, 540, 14))
           Reg.VALOR_NEGOCIADO = Trim(Mid(linha, 554, 14))
           Reg.JUROS_PARCELAMENTO = Trim(Mid(linha, 568, 14))
           Reg.VALOR_ENTRADA = Trim(Mid(linha, 582, 14))
           Reg.QTD_PARCELAS = Trim(Mid(linha, 596, 2))
           Reg.VALOR_PARCELAS = Trim(Mid(linha, 598, 14))

           For i1 = 1 To 15
               If i1 = 1 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 612, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 623, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 633, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 647, 14))
               End If
               If i1 = 2 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 661, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 672, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 682, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 696, 14))
               End If
               If i1 = 3 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 710, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 721, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 731, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 745, 14))
               End If
               If i1 = 4 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 759, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 770, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 780, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 794, 14))
               End If
               If i1 = 5 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 808, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 819, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 829, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 843, 14))
               End If
               If i1 = 6 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 857, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 868, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 878, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 892, 14))
               End If
               If i1 = 7 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 906, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 917, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 927, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 941, 14))
               End If
               If i1 = 8 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 955, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 966, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 976, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 990, 14))
               End If
               If i1 = 9 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1004, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1015, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 1025, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 1039, 14))
               End If
               If i1 = 10 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1053, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1064, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 1074, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 1088, 14))
               End If
               If i1 = 11 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1102, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1113, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 1123, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 1137, 14))
               End If
               If i1 = 12 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1151, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1162, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 1172, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 1186, 14))
               End If
               If i1 = 13 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1200, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1211, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 1221, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 1235, 14))
               End If
               If i1 = 14 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1249, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1260, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 1270, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 1284, 14))
               End If
               If i1 = 15 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1298, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1309, 10))
                  Reg.VALOR_PRESTACAO(i1) = Trim(Mid(linha, 1319, 14))
                  Reg.VALOR_TOTAL_CONTA(i1) = Trim(Mid(linha, 1333, 14))
               End If
           Next
           Reg.TOTAL = Trim(Mid(linha, 1347, 14))
           Reg.VALOR = Trim(Mid(linha, 1361, 14))
           
           MontaExtrato
           LimpaReg
           
           totRegs = totRegs + 1
           frmPrincipal.txQtdReg.Text = totRegs

           'If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

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
        Open App.Path & "\Aviso_De_Parcelamento_Em_Atraso_Compesa_Int.ps" For Input As #200
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
        Open App.Path & "\Aviso_De_Parcelamento_Em_Atraso_Compesa_Ext.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Aviso de Parcelamento em Atraso Compesa A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer

        Reg.GRUPO_EMPREITEIRA = ""
        Reg.INSCRICAO = ""
        Reg.ESCRITORIO = ""
        Reg.NOME = ""
        Reg.ENDERECO = ""
        Reg.BAIRRO = ""
        Reg.CIDADE = ""
        Reg.UF = ""
        Reg.CEP = ""
        Reg.SEQ = ""
        Reg.MATRICULA = ""
        Reg.SEQUENCIAL = ""
        Reg.ORDEM = ""
        Reg.LINHA_DIGITAVEL = ""
        Reg.BAR_CODE = ""
        Reg.DATA_PARCELAMENTO = ""
        Reg.VALOR_DEBITO_ORIGINAL = ""
        Reg.VALOR_ENCARGOS = ""
        Reg.DESCONTO_CONCEDIDO = ""
        Reg.VALOR_NEGOCIADO = ""
        Reg.JUROS_PARCELAMENTO = ""
        Reg.VALOR_ENTRADA = ""
        Reg.QTD_PARCELAS = ""
        Reg.VALOR_PARCELAS = ""
        Reg.TOTAL = ""
        Reg.VENCIMENTO_BOLETO = ""
        Reg.VALOR = ""
        For i = 1 To 15
            Reg.FATURA(i) = ""
            Reg.VENCIMENTO(i) = ""
            Reg.VALOR_PRESTACAO(i) = ""
            Reg.VALOR_TOTAL_CONTA(i) = ""
        Next
End Sub
