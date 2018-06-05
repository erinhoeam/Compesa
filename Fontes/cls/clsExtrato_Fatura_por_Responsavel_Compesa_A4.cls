VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExt_Fat_Resp"
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

Const AppName = "Compesa", AppDesc = "Extrato Fatura por Responsável Compesa A4"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        NOME As String
        COD_RESP As String
        FATURA As String
        ENDERECO As String
        SEQUENCIAL As String
        TIPO_RESP As String
        QTD_FATURAS As String
        MATRICULA_BOX1() As String
        CONSUMO_BOX1() As String
        VALOR_BOX1() As String
        MATRICULA_BOX2() As String
        CONSUMO_BOX2() As String
        VALOR_BOX2() As String
        EMISSAO As String
        VENCIMENTO As String
        VALOR_TOTAL As String
        NAO_RECEBER_APOS As String
        LINHA_DIGITAVEL As String
        CODE_BAR As String
        PAGINA As String
        SEQ As String
End Type
'Variáveis gerais para os tipos
Private Reg As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim posY As Long
        Dim posX As Long

        posX = 170
        posY = 485

        wRec = wRec + 1
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 LANDSCAPE

       'Nome, Cod Responsavel, Fatura
        Print #outFN, "/Helvetica-Condensed-AP findfont 9 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg.COD_RESP & ") show"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Reg.FATURA & ") show"
        posY = posY + 90
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Mid(Reg.ENDERECO, 1, 60) & ") show"
        posY = posY + 35
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Mid(Reg.ENDERECO, 61, 60) & ") show"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 100
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.TIPO_RESP & ") show"
        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.QTD_FATURAS & ") show"
        posY = posY + 150

        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
        For i = 1 To wQtdFaturas
            Print #outFN, posX + 120 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA_BOX1(i) & ") show"
            Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.CONSUMO_BOX1(i))) & Reg.CONSUMO_BOX1(i) & ") show"
            Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Space(25 - Len(Reg.VALOR_BOX1(i))) & Reg.VALOR_BOX1(i) & ") show"
            
            Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA_BOX2(i) & ") show"
            Print #outFN, posX + 1270 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.CONSUMO_BOX2(i))) & Reg.CONSUMO_BOX2(i) & ") show"
            Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Space(25 - Len(Reg.VALOR_BOX2(i))) & Reg.VALOR_BOX2(i) & ") show"
            
            posY = posY + 28

        Next
       'Total Débito
        posY = 2345
        Print #outFN, "/Helvetica-Condensed-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 180 & " pontos -" & posY & " pontos moveto (" & Reg.EMISSAO & ") show"
        Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO & ") show"
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.VALOR_TOTAL)) & Reg.VALOR_TOTAL & ") show"
        posY = posY + 250
        Print #outFN, "/Helvetica-Condensed-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 1900 & " pontos -" & posY & " pontos moveto (" & Reg.SEQ & ") show"
        posY = posY + 100
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 740 & " pontos -" & posY & " pontos moveto (" & Reg.LINHA_DIGITAVEL & ") show"
        posY = posY + 170
        Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
        Print #outFN, posX + 610 & " pontos -" & posY & " pontos moveto (" & Reg.CODE_BAR & ") show"
        posY = posY + 120
        Print #outFN, "/Helvetica-Condensed-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (" & Reg.NAO_RECEBER_APOS & ") show"
        Print #outFN, posX + 490 & " pontos -" & posY & " pontos moveto (" & Reg.COD_RESP & ") show"
        Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.FATURA & ") show"
        posY = posY + 100
        Print #outFN, posX + 180 & " pontos -" & posY & " pontos moveto (" & Reg.EMISSAO & ") show"
        Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO & ") show"
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.VALOR_TOTAL)) & Reg.VALOR_TOTAL & ") show"
        posY = posY + 330
        Print #outFN, "/Helvetica-Condensed-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Reg.PAGINA & ") show"
        Print #outFN, posX + 1900 & " pontos -" & posY & " pontos moveto (" & Reg.SEQ & ") show"
        posY = posY + 70
        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "
        'If wRec = 10 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wContLinha As Integer, wFlagFaturas As Boolean, wCont1 As Integer, wFlagDetalhes As Boolean, wParte As Integer
        wFlagFaturas = False
        wFlagDetalhes = False
        wQtdFaturas = 0
        wContLinha = 0
        totRegs = 0
        wCont1 = 0
        wParte = 1

        'Dividir Registros
        Open pastaProcessar & "temporario.tmp" For Output As #24
        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha

           If Trim(Mid(linha, 1, 2)) = "1" Then
              wCont1 = wCont1 + 1
           End If
           
           If wCont1 = 3 Then
              Print #24, ""
              Print #24, linha
              wCont1 = 1
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
       ' frmPrincipal.ProgressBarProcess.Max = LOF(29)
       ' frmPrincipal.ProgressBarProcess.Value = 0.0001

        Do While Not EOF(29)
           DoEvents
           Line Input #29, linha

           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")
           wContLinha = wContLinha + 1

           Select Case wParte
                Case 1
                    Select Case wContLinha
                         Case 2
                             Reg.NOME = Trim(Mid(linha, 3, 50))
                         Case 3
                             Reg.COD_RESP = Trim(Mid(linha, 58, 15))
                             Reg.FATURA = Trim(Mid(linha, 92, 7))
                         Case 4
                             Reg.ENDERECO = Trim(Mid(linha, 3, 120))
                         Case 5
                             Reg.SEQUENCIAL = Trim(Mid(linha, 90, 10))
                         Case 7
                             Reg.TIPO_RESP = Trim(Mid(linha, 3, 50))
                             Reg.QTD_FATURAS = Trim(Mid(linha, 95, 4))
                         Case 10
                             wFlagDetalhes = True
                     End Select
                Case 2
                    Select Case wContLinha
                         Case 3
                             Reg.SEQ = Trim(Mid(linha, 103, 7))
                         Case 4
                             Reg.LINHA_DIGITAVEL = Trim(Mid(linha, 26, 55))
                         Case 6
                             Reg.CODE_BAR = Trim(Mid(linha, 33, 112))
                         Case 8
                             Reg.NAO_RECEBER_APOS = Trim(Mid(linha, 7, 10))
                         Case 13
                             Reg.PAGINA = Trim(Mid(linha, 97, 13))
                     End Select
            End Select
            
            If wFlagDetalhes Then
               If Mid(linha, 1, 2) = "10" Then
                  wContLinha = 0
                  wParte = 2
                  wFlagDetalhes = False
                  
                  Reg.EMISSAO = Trim(Mid(linha, 11, 10))
                  Reg.VENCIMENTO = Trim(Mid(linha, 45, 10))
                  Reg.VALOR_TOTAL = Trim(Mid(linha, 82, 15))
               End If
            End If
            
            If wFlagDetalhes Then
               wQtdFaturas = wQtdFaturas + 1
               ReDim Preserve Reg.MATRICULA_BOX1(wQtdFaturas)
               ReDim Preserve Reg.CONSUMO_BOX1(wQtdFaturas)
               ReDim Preserve Reg.VALOR_BOX1(wQtdFaturas)
               
               ReDim Preserve Reg.MATRICULA_BOX2(wQtdFaturas)
               ReDim Preserve Reg.CONSUMO_BOX2(wQtdFaturas)
               ReDim Preserve Reg.VALOR_BOX2(wQtdFaturas)
               
               Reg.MATRICULA_BOX1(wQtdFaturas) = Trim(Mid(linha, 3, 9))
               Reg.CONSUMO_BOX1(wQtdFaturas) = Trim(Mid(linha, 14, 14))
               Reg.VALOR_BOX1(wQtdFaturas) = Trim(Mid(linha, 33, 20))
               
               Reg.MATRICULA_BOX2(wQtdFaturas) = Trim(Mid(linha, 70, 9))
               Reg.CONSUMO_BOX2(wQtdFaturas) = Trim(Mid(linha, 82, 14))
               Reg.VALOR_BOX2(wQtdFaturas) = Trim(Mid(linha, 100, 20))
            End If

            If Mid(linha, 1, 1) = "" Then
               wContLinha = 0
               wFlagDetalhes = False
               wParte = 1
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
        Open App.Path & "\Extrato_Faturas_Responsavel_Compesa_A4.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Extrato Fatura por Responsável Compesa A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Reg.NOME = ""
        Reg.COD_RESP = ""
        Reg.FATURA = ""
        Reg.ENDERECO = ""
        Reg.SEQUENCIAL = ""
        Reg.TIPO_RESP = ""
        Reg.QTD_FATURAS = ""
        Reg.EMISSAO = ""
        Reg.VENCIMENTO = ""
        Reg.VALOR_TOTAL = ""
        Reg.NAO_RECEBER_APOS = ""
        Reg.LINHA_DIGITAVEL = ""
        Reg.CODE_BAR = ""
        Reg.PAGINA = ""
        ReDim Reg.MATRICULA_BOX1(0) As String
        ReDim Reg.CONSUMO_BOX1(0) As String
        ReDim Reg.VALOR_BOX1(0) As String
        ReDim Reg.MATRICULA_BOX2(0) As String
        ReDim Reg.CONSUMO_BOX2(0) As String
        ReDim Reg.VALOR_BOX2(0) As String
        Reg.SEQ = ""
End Sub
