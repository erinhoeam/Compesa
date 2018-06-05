VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdemDeSupressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap Consultoria
'                 Analista de Sistemas
'Data: 20 de Agosto de 2008
'Contatos - Email: ericsson.alves@centauro-ne.com.br
'                  erinhoeam@hotmail.com
'                  aziel@sistemap.com.br
'           Fone: (81) 9832.6084
'Aplicação: Aviso de Corte Adminsitrativo Compesa A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double
Dim wQtdFaturas(2) As Long
Dim wIndiceFaturas As Long

Const AppName = "CAERN", AppDesc = "Aviso de Corte Administrativo Compesa A4/A5"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        ESCRITORIO As String
        CICLO As String
        INSCRICAO As String
        NOME As String
        MATRICULA As String
        ENDERECO As String
        VALIDO As String
        NR_ORDEM As String
        GRUPO As String

        A_E As String
        ECONOMIAS As String
        RES As String
        COM As String
        IND As String
        PUB As String
        TIPO_CONSUMIDOR As String
        N_HIDROMETRO As String
        DATA_EMISSAO As String
        NAO_RECEBER_APOS As String
        FATURA_EM_ATRASO() As String
        VENCIMENTO() As String
        VALOR() As String
        TOTAL_DEBITO As String
        ESCRITORIO2 As String

        SEQUENCIA As String
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim posY As Long
        Dim posX As Long
        Dim i2 As Integer

        wRec = wRec + 1
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        For i = 1 To 2
            totRegs = totRegs + 1
            Select Case i
                Case 1
                  posX = 100
                  posY = 215
                Case 2
                  posX = 1800
                  posY = 215
            End Select
            'Escritorio, Ciclo
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 710 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESCRITORIO & ") show"
            Print #outFN, posX + 1150 & " pontos -" & posY & " pontos moveto (" & Reg(i).CICLO & ") show"
            posY = posY + 90
            'Inscrição, Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 650 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            posY = posY + 80
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            'Endereço, Grupo, Numero de Ordem
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
            Print #outFN, posX + 1170 & " pontos -" & posY & " pontos moveto (" & Reg(i).VALIDO & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg(i).NR_ORDEM & ") show"
            posY = posY + 95
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            'Com, ind, Pub, Tipo Consumidor, Data Emissão
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).RES & ") show"
            Print #outFN, posX + 120 & " pontos -" & posY & " pontos moveto (" & Reg(i).COM & ") show"
            Print #outFN, posX + 225 & " pontos -" & posY & " pontos moveto (" & Reg(i).IND & ") show"
            Print #outFN, posX + 325 & " pontos -" & posY & " pontos moveto (" & Reg(i).PUB & ") show"
            Print #outFN, posX + 530 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_EMISSAO & ") show"
            Print #outFN, posX + 750 & " pontos -" & posY & " pontos moveto (" & Reg(i).TIPO_CONSUMIDOR & ") show"
            Print #outFN, posX + 1000 & " pontos -" & posY & " pontos moveto (" & Reg(i).N_HIDROMETRO & ") show"
            Print #outFN, posX + 1340 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
            Print #outFN, posX + 1480 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQUENCIA & ") show"
            'Faturas em Atraso
            posY = posY + 140
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 8 scalefont setfont"
            For i2 = 1 To wQtdFaturas(i)
                Print #outFN, posX - 50 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).FATURA_EM_ATRASO(i2))) & Reg(i).FATURA_EM_ATRASO(i2) & ") show"
                Print #outFN, posX + 220 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).VENCIMENTO(i2))) & Reg(i).VENCIMENTO(i2) & ") show"

                posY = posY + 30

            Next
            'Total Débito
            Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
            posY = 1100
            Print #outFN, posX + 110 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg(i).TOTAL_DEBITO)) & Reg(i).TOTAL_DEBITO & ") show"
            posY = posY + 350
            'Escritorio, Sequencia
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESCRITORIO2 & ") show"
            posY = posY + 530
            'Inscrição, Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 580 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            posY = posY + 95
            'Endereço, Grupo, Numero de Ordem
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg(i).NR_ORDEM & ") show"
            posY = posY + 80
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            'Com, ind, Pub, Tipo Consumidor, Data Emissão
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).RES & ") show"
            Print #outFN, posX + 120 & " pontos -" & posY & " pontos moveto (" & Reg(i).COM & ") show"
            Print #outFN, posX + 225 & " pontos -" & posY & " pontos moveto (" & Reg(i).IND & ") show"
            Print #outFN, posX + 325 & " pontos -" & posY & " pontos moveto (" & Reg(i).PUB & ") show"
            Print #outFN, posX + 440 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_EMISSAO & ") show"
            Print #outFN, posX + 670 & " pontos -" & posY & " pontos moveto (" & Reg(i).TIPO_CONSUMIDOR & ") show"
            Print #outFN, posX + 1000 & " pontos -" & posY & " pontos moveto (" & Reg(i).A_E & ") show"
            Print #outFN, posX + 1170 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
            Print #outFN, posX + 1400 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQUENCIA & ") show"
            posY = posY + 290
            'Sequencia e Recuperação
            Print #outFN, "/Helvetica-AP findfont 4.5 scalefont setfont"
            Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "000000") & ") show"
            Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"

        Next
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "
'        If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wCont As Integer, wIndice As Integer, wCont11 As Integer, wContLinha As Integer
        wCont = 0
        wIndice = 1
        wQtdFaturas(1) = 0
        wQtdFaturas(2) = 0
        wContLinha = 0
        wIndiceFaturas = 0
        totRegs = 0

        'Dividir Registros
        Open pastaProcessar & "temporario.tmp" For Output As #24
        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha

           If Mid(linha, 1, 2) = "11" Then
              wCont = wCont + 1
           End If
           If wCont = 4 Then
              wCont = 1
              Print #24, ""
              Print #24, linha
           Else
              Print #24, linha
           End If

           'If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        Close #24
        Close #inFN

        totRegs = 0
        wCont = 1

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

           If Mid(linha, 1, 2) = "11" Then
              wCont11 = wCont11 + 1
           End If

           Select Case wCont
                Case 1
                    Select Case wContLinha
                        Case 1
                            Reg(wIndice).ESCRITORIO = Trim(Mid(linha, 3, 22))
                            Reg(wIndice).CICLO = Trim(Mid(linha, 31, 7))
                        Case 2
                            Reg(wIndice).INSCRICAO = Trim(Mid(linha, 3, 25))
                            Reg(wIndice).NOME = Trim(Mid(linha, 38, 60))
                        Case 3
                            Reg(wIndice).MATRICULA = Trim(Mid(linha, 69, 10))
                        Case 4
                            Reg(wIndice).ENDERECO = Trim(Mid(linha, 3, 64))
                            Reg(wIndice).VALIDO = Trim(Mid(linha, 67, 10))
                        Case 5
                            Reg(wIndice).GRUPO = Trim(Mid(linha, 70, 4))
                            Reg(wIndice).NR_ORDEM = Trim(Mid(linha, 71, 10))
                        Case 6
                            Reg(wIndice).RES = Mid(linha, 4, 3)
                            Reg(wIndice).COM = Mid(linha, 10, 3)
                            Reg(wIndice).IND = Mid(linha, 16, 3)
                            Reg(wIndice).PUB = Mid(linha, 22, 3)
                            Reg(wIndice).DATA_EMISSAO = Trim(Mid(linha, 32, 10))
                            Reg(wIndice).TIPO_CONSUMIDOR = Trim(Mid(linha, 46, 14))
                            Reg(wIndice).N_HIDROMETRO = Trim(Mid(linha, 60, 10))
                            Reg(wIndice).GRUPO = Trim(Mid(linha, 71, 8))
                            Reg(wIndice).SEQUENCIA = Trim(Mid(linha, 85, 7))
                    End Select
                    If Mid(linha, 1, 2) <> "12" Then
                       If wCont11 = 2 Then
                          wQtdFaturas(wIndice) = wQtdFaturas(wIndice) + 1
                          ReDim Preserve Reg(wIndice).FATURA_EM_ATRASO(wQtdFaturas(wIndice))
                          ReDim Preserve Reg(wIndice).VENCIMENTO(wQtdFaturas(wIndice))

                          Reg(wIndice).FATURA_EM_ATRASO(wQtdFaturas(wIndice)) = Trim(Mid(linha, 4, 10))
                          Reg(wIndice).VENCIMENTO(wQtdFaturas(wIndice)) = Trim(Mid(linha, 16, 10))
                       End If
                    End If
                    If Mid(linha, 1, 2) = "12" Then
                       Reg(wIndice).TOTAL_DEBITO = Trim(Mid(linha, 4, 22))
                       wCont = wCont + 1
                       wContLinha = 0
                    End If
               Case 2
                    Select Case wContLinha
                        Case 3
                            Reg(wIndice).ESCRITORIO2 = Trim(Mid(linha, 3, 90))
                        Case 8
                            Reg(wIndice).A_E = Trim(Mid(linha, 60, 4))
                    End Select
           End Select

           If Mid(linha, 1, 1) = "" Then
              wIndice = wIndice + 1
              wCont = 1
              wCont11 = 0
              wContLinha = 0
           End If

           If wIndice = 3 Then
              wCont = 1
              wCont11 = 0
              wContLinha = 0
              wIndice = 1
              MontaExtrato
              frmPrincipal.txQtdReg.Text = totRegs
              LimpaReg
              wQtdFaturas(1) = 0
              wQtdFaturas(2) = 0
           End If

         '  If Seek(29) <= frmPrincipal.ProgressBarProcess.Max Then _
         '     frmPrincipal.ProgressBarProcess.Value = Seek(29)
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
        Open App.Path & "\Ordem_Supressao.ps" For Input As #200
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
'        Print #outFN, "/DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
'        Print #outFN, ">> /ReusableStreamDecode filter"
'        Open App.Path & "\INT_Contra_Cheque.ps" For Input As #200
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
        Print #outFN, "%Aplicação: AVISO DE CORTE ADMINISTRATIVO COMPESA A4/A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        For i = 1 To 2
            'Reg(i).ESCRITORIO = ""
            'Reg(i).CICLO = ""
            Reg(i).INSCRICAO = ""
            Reg(i).NOME = ""
            Reg(i).MATRICULA = ""
            Reg(i).ENDERECO = ""
            Reg(i).VALIDO = ""
            Reg(i).NR_ORDEM = ""
            Reg(i).GRUPO = ""

            Reg(i).A_E = ""
            Reg(i).ECONOMIAS = ""
            Reg(i).RES = ""
            Reg(i).COM = ""
            Reg(i).IND = ""
            Reg(i).PUB = ""
            Reg(i).TIPO_CONSUMIDOR = ""
            Reg(i).DATA_EMISSAO = ""
            Reg(i).NAO_RECEBER_APOS = ""
            ReDim FATURA_EM_ATRASO(0) As String
            ReDim VENCIMENTO(0) As String
            Reg(i).TOTAL_DEBITO = ""
            Reg(i).ESCRITORIO2 = ""

            Reg(i).SEQUENCIA = ""
        Next
End Sub


