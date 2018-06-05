VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConta_Compesa_A5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Ericsson Alves de Melo
'                 Analista de Sistemas
'Data: 27 de Novembro de 2008
'Contatos - Email: ericsson@sistemap.com.br
'                  erinhoeam@hotmail.com
'           Fone: (81) 9985.9131
'Aplicação: Conta Compesa A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double
Dim wQtdFaturas(2) As Long
Dim wIndiceFaturas As Long

Const AppName = "COMPESA", AppDesc = "Conta Compesa A4/A5"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        TIPO_CONTA As String
        ESCRITORIO As String
        MATRICULA As String
        CLIENTE As String
        VENCIMENTO As String
        INSCRICAO As String
        ENDERECO1 As String
        ENDERECO2 As String
        FATURA As String
        RESPONSAVEL As String
        ENDERECO_ENTREGA As String
        ENDERECO_ENTREGA1 As String
        CEP As String
        AGUA As String
        ESGOTO As String
        ULTIMOS_CONSUMOS(3) As String
        LEITURA_ANTERIOR As String
        PERIODO As String
        CONSUMO As String
        CONSUMO2 As String
        ECONOMIA As String
        CONSUMO_POR_ECONOMIA As String
        CONSUMO_POR_RATEIO As String
        COD_AUXILIAR As String
        DESCRICAO() As String
        FAIXA() As String
        VALOR() As String
        TOTAL As String
        MENSAGEM(3) As String
        GERENCIA_REGIONAL As String
        MES_ANO As String
        TURBIDEZ As String
        CLORO As String
        SEQ As String
        LINHA_DIGITAVEL As String
        BAR_CODE As String
        NAO_RECEBER As String
        GRUPO As String
        FIRMA As String
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim i2 As Long
        Dim i3 As Long
        Dim i4 As Long
        Dim posY As Long
        Dim posX As Long
        
        wRec = wRec + 1
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        For i = 1 To 2
            If Len(Trim(Reg(i).CLIENTE)) > 0 Then totRegs = totRegs + 1
            Select Case i
                Case 1
                  posX = 100
                  posY = 260
                Case 2
                  posX = 1860
                  posY = 260
            End Select
            'Rota
            Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
            Print #outFN, posX + 300 & " pontos -" & posY - 190 & " pontos moveto (" & Reg(i).TIPO_CONTA & ") show"
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Escritório: " & Reg(i).ESCRITORIO & ") show"
            posY = posY + 60
            'Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 350 & " pontos -" & posY & " pontos moveto (" & Reg(i).CLIENTE & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 50 & " pontos -" & posY - 10 & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            Print #outFN, posX + 1300 & " pontos -" & posY - 10 & " pontos moveto (" & Reg(i).VENCIMENTO & ") show"
            posY = posY + 90
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 20 & " pontos -" & posY + 20 & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO1 & ") show"
            Print #outFN, posX + 1370 & " pontos -" & posY + 20 & " pontos moveto (" & Reg(i).FATURA & ") show"
            Print #outFN, posX + 400 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).ENDERECO2 & ") show"
            posY = posY + 100
            Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
            Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).RESPONSAVEL & ") show"
            Print #outFN, posX + 180 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO_ENTREGA & ") show"
            Print #outFN, posX + 180 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).ENDERECO_ENTREGA1 & ") show"
            Print #outFN, posX + 1015 & " pontos -" & posY & " pontos moveto (" & Reg(i).AGUA & ") show"
            Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESGOTO & ") show"
            posY = posY + 110
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
            For i2 = 1 To 3
                Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).ULTIMOS_CONSUMOS(i2) & ") show"
                If i2 = 2 Then
                   Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg(i).LEITURA_ANTERIOR & ") show"
                   Print #outFN, posX + 830 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).PERIODO & ") show"
                End If
                posY = posY + 30
            Next
            Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO & ") show"
            Print #outFN, posX + 1180 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO2 & ") show"
            posY = posY + 30
            Print #outFN, posX + 100 & " pontos -" & posY & " pontos moveto (" & Reg(i).ECONOMIA & ") show"
            Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO_POR_ECONOMIA & ") show"
            Print #outFN, posX + 380 & " pontos -" & posY & " pontos moveto (" & Reg(i).COD_AUXILIAR & ") show"
            Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO_POR_RATEIO & ") show"
            posY = posY + 120
            For i3 = 1 To wQtdFaturas(i)
                Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).DESCRICAO(i3) & ") show"
                Print #outFN, posX + 1100 & " pontos -" & posY & " pontos moveto (" & Reg(i).FAIXA(i3) & ") show"
                Print #outFN, posX + 1040 & " pontos -" & posY & " pontos moveto (" & Space(30 - Len(Reg(i).VALOR(i3))) & Reg(i).VALOR(i3) & ") show"

                posY = posY + 30
            Next
            posY = 1350
'            'Total Débito
            Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Space(30 - Len(Reg(i).TOTAL)) & Reg(i).TOTAL & ") show"
            posY = posY + 70
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
            For i4 = 1 To 3
                Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).MENSAGEM(i4) & ") show"
                
                posY = posY + 30
            Next
            posY = posY + 100
            Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
            Print #outFN, posX + 240 & " pontos -" & posY & " pontos moveto (" & Reg(i).GERENCIA_REGIONAL & ") show"
            Print #outFN, posX + 1100 & " pontos -" & posY & " pontos moveto (" & Reg(i).MES_ANO & ") show"
            posY = posY + 70
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg(i).TURBIDEZ & ") show"
            Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (" & Reg(i).CLORO & ") show"
            posY = posY + 70
            Print #outFN, posX + 1400 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQ & ") show"
            posY = posY + 100
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg(i).LINHA_DIGITAVEL & ") show"
            posY = posY + 180
            Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
            Print #outFN, posX + 280 & " pontos -" & posY & " pontos moveto (" & Reg(i).BAR_CODE & ") show"
            posY = posY + 100
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 420 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 1000 & " pontos -" & posY & " pontos moveto (" & Reg(i).FATURA & ") show"
            Print #outFN, posX + 1400 & " pontos -" & posY & " pontos moveto (" & Reg(i).NAO_RECEBER & ") show"
            posY = posY + 20
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            posY = posY + 115
            Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Space(30 - Len(Reg(i).TOTAL)) & Reg(i).TOTAL & ") show"
            posY = posY + 20
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
            Print #outFN, posX + 250 & " pontos -" & posY & " pontos moveto (" & Reg(i).VENCIMENTO & ") show"
            posY = posY + 60
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 250 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
            Print #outFN, posX + 750 & " pontos -" & posY & " pontos moveto (" & Reg(i).FIRMA & ") show"
            Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (Seq: " & Reg(i).SEQ & ") show"
            Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
            posY = posY + 75
            Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
            Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
            Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (OHNIRE) show"
        Next
        Print #outFN, "showpage"
        'Fim da Parte Interna
        '**********Inicio Parte Externa************************************************
        'Print #outFN, "save EPSForm2 execform restore"
        'Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        'Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

'       If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wCont As Long, wIndice As Integer, wIndiceConsumo As Integer, wIndiceMensagem As Integer, wContRegs As Integer
        Dim wFlagDetalhe As Boolean
        wRec = 0
        wCont = 0
        wIndice = 1
        wQtdFaturas(1) = 0
        wQtdFaturas(2) = 0
        wIndiceConsumo = 0
        wIndiceMensagem = 0
        wContRegs = 1
        wFlagDetalhe = False
        totRegs = 0

        'Dividir Registros
        Open pastaProcessar & "temporario.tmp" For Output As #24
        frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha

           If Mid(linha, 1, 2) = "-1" Then
              wCont = wCont + 1
           End If
           If wCont = 2 Then
              wCont = 1
              Print #24, linha
              Print #24, ""
              wCont = 0
           Else
              Print #24, linha
           End If

           If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
              frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        Close #24
        Close #inFN

        totRegs = 0
        wCont = 0

        'Gravar Cabeçalho
        GravaHeader

        wArq = frmPrincipal.LstSpool.List(frmPrincipal.LstSpool.ListIndex)
        Open pastaProcessar & "temporario.tmp" For Input As #29
        frmPrincipal.ProgressBarProcess.Max = LOF(29)
        frmPrincipal.ProgressBarProcess.Value = 0.0001

        Do While Not EOF(29)
           DoEvents
           Line Input #29, linha
           
           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")
           
           Select Case wContRegs
                  Case 1
                       wCont = wCont + 1

                       Select Case wCont
                           Case 1
                                Reg(wIndice).TIPO_CONTA = Trim(Mid(linha, 3, 90))
                           Case 2
                                Reg(wIndice).ESCRITORIO = Trim(Mid(linha, 3, 72))
                           Case 3
                                Reg(wIndice).MATRICULA = Trim(Mid(linha, 3, 10))
                                Reg(wIndice).VENCIMENTO = Trim(Mid(linha, 65, 10))
                           Case 4
                                Reg(wIndice).CLIENTE = Trim(Mid(linha, 3, 86))
                           Case 5
                                Reg(wIndice).ENDERECO1 = Trim(Mid(linha, 3, 77))
                                Reg(wIndice).FATURA = Trim(Mid(linha, 80, 9))
                           Case 6
                                Reg(wIndice).ENDERECO2 = Trim(Mid(linha, 3, 77))
                           Case 7
                                Reg(wIndice).INSCRICAO = Trim(Mid(linha, 3, 20))
                           Case 8
                                Reg(wIndice).RESPONSAVEL = Trim(Mid(linha, 3, 7))
                                Reg(wIndice).ENDERECO_ENTREGA = Trim(Mid(linha, 10, 59))
                                
                                Reg(wIndice).AGUA = Trim(Mid(linha, 69, 16))
                                Reg(wIndice).ESGOTO = Trim(Mid(linha, 85, 16))
                           Case 9
                                'Reg(wIndice).CEP = Trim(Mid(linha, 3, 20))
                                Reg(wIndice).ENDERECO_ENTREGA1 = Trim(Mid(linha, 3, 56))
                           Case 11 To 13
                                wIndiceConsumo = wIndiceConsumo + 1
                                Reg(wIndice).ULTIMOS_CONSUMOS(wIndiceConsumo) = Trim(Mid(linha, 3, 38))
                                If wCont = 11 Then
                                   Reg(wIndice).LEITURA_ANTERIOR = Mid(linha, 45, 40)
                                ElseIf wCont = 12 Then
                                   Reg(wIndice).PERIODO = Trim(Mid(linha, 45, 40))
                                End If
                           Case 14
                                Reg(wIndice).CONSUMO = Trim(Mid(linha, 45, 29))
                                Reg(wIndice).CONSUMO2 = Trim(Mid(linha, 74, 29))
                           Case 15
                                Reg(wIndice).ECONOMIA = Trim(Mid(linha, 3, 6))
                                Reg(wIndice).CONSUMO_POR_ECONOMIA = Trim(Mid(linha, 9, 12))
                                Reg(wIndice).COD_AUXILIAR = Trim(Mid(linha, 26, 20))
                                Reg(wIndice).CONSUMO_POR_RATEIO = Trim(Mid(linha, 48, 40))
                           Case 18
                                wFlagDetalhe = True
                       End Select
                       If Mid(linha, 1, 2) = "12" Then
                          wFlagDetalhe = False
                          Reg(wIndice).TOTAL = Trim(Mid(linha, 3, 84))
                          wContRegs = wContRegs + 1
                          wCont = 0
                       End If
                       If wFlagDetalhe Then
                          wQtdFaturas(wIndice) = wQtdFaturas(wIndice) + 1
                          ReDim Preserve Reg(wIndice).DESCRICAO(wQtdFaturas(wIndice))
                          ReDim Preserve Reg(wIndice).FAIXA(wQtdFaturas(wIndice))
                          ReDim Preserve Reg(wIndice).VALOR(wQtdFaturas(wIndice))
                          
                          Reg(wIndice).DESCRICAO(wQtdFaturas(wIndice)) = Mid(linha, 3, 50)
                          Reg(wIndice).FAIXA(wQtdFaturas(wIndice)) = Trim(Mid(linha, 53, 15))
                          Reg(wIndice).VALOR(wQtdFaturas(wIndice)) = Trim(Mid(linha, 68, 20))
                       End If
                  Case 2
                       wCont = wCont + 1

                       Select Case wCont
                           Case 1 To 3
                                wIndiceMensagem = wIndiceMensagem + 1
                                Reg(wIndice).MENSAGEM(wIndiceMensagem) = Trim(Mid(linha, 3, 85))
                           Case 4
                                Reg(wIndice).GERENCIA_REGIONAL = Trim(Mid(linha, 3, 50))
                                Reg(wIndice).MES_ANO = Trim(Mid(linha, 53, 7))
                           Case 5
                                Reg(wIndice).TURBIDEZ = Trim(Mid(linha, 3, 24))
                                Reg(wIndice).CLORO = Trim(Mid(linha, 27, 24))
                           Case 6
                                Reg(wIndice).SEQ = Trim(Mid(linha, 3, 74))
                           Case 7
                                Reg(wIndice).LINHA_DIGITAVEL = Trim(Mid(linha, 3, 86))
                           Case 9
                                Reg(wIndice).BAR_CODE = Trim(Mid(linha, 3, 129))
                           Case 11
                                Reg(wIndice).NAO_RECEBER = Trim(Mid(linha, 83, 10))
                           Case 14
                                Reg(wIndice).GRUPO = Trim(Mid(linha, 3, 20))
                                Reg(wIndice).FIRMA = Trim(Mid(linha, 20, 45))
                       End Select
           End Select

           If Mid(linha, 1, 1) = "" Then
              wIndice = wIndice + 1
              wContRegs = 1
              wCont = 0
              wIndiceConsumo = 0
              wIndiceMensagem = 0
              wFlagDetalhe = False
           End If
           
           If wIndice = 3 Then
              wIndice = 1
              wIndiceConsumo = 0
              wIndiceMensagem = 0
              wFlagDetalhe = False
              wContRegs = 1
              wCont = 0
              MontaExtrato
              frmPrincipal.txQtdReg.Text = totRegs
              LimpaReg
              wQtdFaturas(1) = 0
              wQtdFaturas(2) = 0
           End If

           If Seek(29) <= frmPrincipal.ProgressBarProcess.Max Then _
              frmPrincipal.ProgressBarProcess.Value = Seek(29)
        Loop
        If wIndice > 1 Then MontaExtrato
        Screen.MousePointer = vbDefault

        frmResultados.txtTotCli_STL.Text = totRegs
        frmResultados.txtTotReg_STL.Text = wRec
        frmResultados.txArquivo.Text = wArq

        frmResultados.txtAplicacao = AppDesc

        Close #outFN
        Close #29

        Kill pastaProcessar & "*.tmp"

        frmPrincipal.MousePointer = vbDefault
        frmPrincipal.ProgressBarProcess.Value = 0.0001

   Exit Sub
'--------------------------------------------------------------------------------
Erro:
   Close #outFN
   Close #inFN
   
   frmPrincipal.MousePointer = vbDefault
   frmPrincipal.ProgressBarProcess.Value = 0.0001
   
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
        Open App.Path & "\Conta_Int.ps" For Input As #200
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
'        Print #outFN, "/ImageData2"
'        Print #outFN, "currentfile"
'        Print #outFN, "<< /Filter /SubFileDecode"
'        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
'        Print #outFN, ">> /ReusableStreamDecode filter"
'        Open App.Path & "\Conta_Ext.ps" For Input As #200
'        While Not EOF(200)
'            Line Input #200, wCabeca
'            Print #outFN, wCabeca
'        Wend
'        Close #200
'        Print #outFN, "*EOD*"
'        Print #outFN, "def"
'        Print #outFN, " "
'        Print #outFN, "/EPSForm2"
'        Print #outFN, "<< /FormType 1"
'        Print #outFN, "   /BBox [0 0 595 900]"
'        Print #outFN, "   /Matrix [ 1 0 0 1 0 0]"
'        Print #outFN, "   /PaintProc"
'        Print #outFN, "   { pop"
'        Print #outFN, "       /ostate save def"
'        Print #outFN, "         /showpage {} def"
'        Print #outFN, "         /setpagedevice /pop load def"
'        Print #outFN, "         ImageData2 0 setfileposition ImageData2 cvx exec"
'        Print #outFN, "       ostate restore"
'        Print #outFN, "   } bind"
'        Print #outFN, ">> def"
'        Print #outFN, "%% Fim da Segunda Imagem"
'        Print #outFN, " "
        '-------------------------------------------------------------------
        Print #outFN, "%%============================"
        Print #outFN, "%% Inicio dos Dados Variaveis"
        Print #outFN, "%%============================"
        Print #outFN, " "
        Print #outFN, "%Aplicação: CONTA COMPESA A4/A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        For i = 1 To 2
            Reg(i).TIPO_CONTA = ""
            Reg(i).ESCRITORIO = ""
            Reg(i).MATRICULA = ""
            Reg(i).CLIENTE = ""
            Reg(i).VENCIMENTO = ""
            Reg(i).INSCRICAO = ""
            Reg(i).ENDERECO1 = ""
            Reg(i).ENDERECO2 = ""
            Reg(i).FATURA = ""
            Reg(i).RESPONSAVEL = ""
            Reg(i).ENDERECO_ENTREGA = ""
            Reg(i).CEP = ""
            Reg(i).AGUA = ""
            Reg(i).ESGOTO = ""
            Reg(i).ULTIMOS_CONSUMOS(1) = ""
            Reg(i).ULTIMOS_CONSUMOS(2) = ""
            Reg(i).ULTIMOS_CONSUMOS(3) = ""
            Reg(i).LEITURA_ANTERIOR = ""
            Reg(i).PERIODO = ""
            Reg(i).ECONOMIA = ""
            Reg(i).CONSUMO = ""
            Reg(i).CONSUMO2 = ""
            Reg(i).CONSUMO_POR_ECONOMIA = ""
            Reg(i).COD_AUXILIAR = ""
            ReDim Reg(i).DESCRICAO(0) As String
            ReDim Reg(i).FAIXA(0) As String
            ReDim Reg(i).VALOR(0) As String
            Reg(i).TOTAL = ""
            Reg(i).MENSAGEM(1) = ""
            Reg(i).MENSAGEM(2) = ""
            Reg(i).MENSAGEM(3) = ""
            Reg(i).GERENCIA_REGIONAL = ""
            Reg(i).MES_ANO = ""
            Reg(i).TURBIDEZ = ""
            Reg(i).CLORO = ""
            Reg(i).SEQ = ""
            Reg(i).LINHA_DIGITAVEL = ""
            Reg(i).BAR_CODE = ""
            Reg(i).NAO_RECEBER = ""
            Reg(i).GRUPO = ""
            Reg(i).FIRMA = ""
        Next
End Sub
