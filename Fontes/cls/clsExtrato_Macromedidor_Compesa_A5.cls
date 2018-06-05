VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtrato_Macromedidor_A5"
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
'Aplicação: Extrato Macromedidor Compesa A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Extrato Macromedidor Compesa A4/A5"

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
        INSCRICAO As String
        NOME As String
        MATRICULA As String
        ENDERECO As String
        ENDERECO2 As String
        FATURA As String
        RESPONSAVEL As String
        ENDERECO_ENTREGA As String
        ULTIMOS_CONSUMOS(3) As String
        DATA_LEITURA As String
        LEITURA As String
        CONSUMO_TOTAL As String
        NUMERO_UNIDADES As String
        CONSUMO_UNIDADE As String
        COD_AUXILIAR As String
        SITUACAO_FATURAMENTO(3) As String
        EXTRATO(4) As String
        SEQ As String
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
            If Len(Trim(Reg(i).NOME)) > 0 Then totRegs = totRegs + 1
            Select Case i
                Case 1
                  posX = 100
                  posY = 275
                Case 2
                  posX = 1860
                  posY = 275
            End Select
            'Rota
            Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Escritório: " & Reg(i).ESCRITORIO & ") show"
            posY = posY + 60
            'Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 580 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 1340 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            posY = posY + 80
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 1340 & " pontos -" & posY & " pontos moveto (" & Reg(i).FATURA & ") show"
            posY = posY + 30
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO2 & ") show"
            posY = posY + 70
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            'Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).RESPONSAVEL & ") show"
            posY = posY + 100
            Print #outFN, posX + 850 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_LEITURA & ") show"
            Print #outFN, posX + 1080 & " pontos -" & posY & " pontos moveto (" & Reg(i).LEITURA & ") show"
            Print #outFN, posX + 1340 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO_TOTAL & ") show"
            posY = posY + 5
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
            For i2 = 1 To 3
                Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg(i).ULTIMOS_CONSUMOS(i2) & ") show"
                posY = posY + 30
            Next
            posY = posY - 20
            For i3 = 1 To 3
                Print #outFN, posX + 870 & " pontos -" & posY & " pontos moveto (" & Reg(i).SITUACAO_FATURAMENTO(i3) & ") show"
                posY = posY + 30
            Next
            posY = posY - 20
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg(i).NUMERO_UNIDADES & ") show"
            Print #outFN, posX + 250 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO_UNIDADE & ") show"
            Print #outFN, posX + 460 & " pontos -" & posY & " pontos moveto (" & Reg(i).COD_AUXILIAR & ") show"
            posY = posY + 130
            Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8 scalefont setfont"
            For i4 = 1 To 4
                Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).EXTRATO(i4) & ") show"
                posY = posY + 30
            Next
            posY = posY + 1375
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 180 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
            Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg(i).FIRMA & ") show"
            Print #outFN, posX + 1390 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQ & ") show"
            posY = posY + 55
            Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
            Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
            Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Next
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

'       If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wCont As Integer, wIndice As Integer, wIndiceUltConsumo As Integer, wIndiceExtrato As Integer
        wRec = 0
        wCont = 0
        wIndice = 1
        wIndiceUltConsumo = 0
        totRegs = 0
        wIndiceExtrato = 0

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
           If wCont = 3 Then
              wCont = 1
              Print #24, linha
              Print #24, ""
              wCont = 0
           Else
              Print #24, linha
           End If

         '  If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
         '     frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        Close #24
        Close #inFN

        wCont = 0

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

                wCont = wCont + 1

                Select Case wCont
                       Case 2
                            Reg(wIndice).ESCRITORIO = Trim(Mid(linha, 3, 72))
                       Case 3
                            Reg(wIndice).INSCRICAO = Trim(Mid(linha, 3, 20))
                            Reg(wIndice).NOME = Trim(Mid(linha, 33, 50))
                       Case 4
                            Reg(wIndice).MATRICULA = Trim(Mid(linha, 67, 9))
                       Case 5
                            Reg(wIndice).ENDERECO = Trim(Mid(linha, 3, 50))
                       Case 6
                            Reg(wIndice).FATURA = Trim(Mid(linha, 67, 9))
                       Case 7
                            Reg(wIndice).ENDERECO2 = Trim(Mid(linha, 3, 75))
                       Case 10 To 12
                            wIndiceUltConsumo = wIndiceUltConsumo + 1
                            Reg(wIndice).ULTIMOS_CONSUMOS(wIndiceUltConsumo) = Trim(Mid(linha, 3, 40))
                            If wCont = 10 Then
                               Reg(wIndice).DATA_LEITURA = Trim(Mid(linha, 49, 10))
                               Reg(wIndice).LEITURA = Trim(Mid(linha, 59, 9))
                               Reg(wIndice).CONSUMO_TOTAL = Trim(Mid(linha, 75, 15))
                            ElseIf wCont = 12 Then
                                   Reg(wIndice).SITUACAO_FATURAMENTO(1) = Trim(Mid(linha, 48, 20))
                            End If
                       Case 13
                            Reg(wIndice).SITUACAO_FATURAMENTO(2) = Trim(Mid(linha, 48, 20))
                       Case 14
                            Reg(wIndice).SITUACAO_FATURAMENTO(3) = Trim(Mid(linha, 48, 20))
                            Reg(wIndice).NUMERO_UNIDADES = Trim(Mid(linha, 3, 6))
                            Reg(wIndice).CONSUMO_UNIDADE = Trim(Mid(linha, 9, 12))
                            Reg(wIndice).COD_AUXILIAR = Trim(Mid(linha, 21, 20))
                       Case 16 To 19
                            wIndiceExtrato = wIndiceExtrato + 1
                            Reg(wIndice).EXTRATO(wIndiceExtrato) = Trim(Mid(linha, 3, 60))
                       Case 20
                            Reg(wIndice).GRUPO = Trim(Mid(linha, 5, 9))
                            Reg(wIndice).FIRMA = Trim(Mid(linha, 37, 20))
                            Reg(wIndice).SEQ = Trim(Mid(linha, 70, 7))
                End Select

                If Mid(linha, 1, 1) = "" Then
                   wIndice = wIndice + 1
                   wIndiceExtrato = 0
                   wIndiceUltConsumo = 0
                   wCont = 0
                End If

                If wIndice = 3 Then
                   wIndice = 1
                   wIndiceExtrato = 0
                   wIndiceUltConsumo = 0
                   wCont = 0
                   MontaExtrato
                   frmPrincipal.txQtdReg.Text = totRegs
                   LimpaReg
                End If

         '       If Seek(29) <= frmPrincipal.ProgressBarProcess.Max Then _
         '          frmPrincipal.ProgressBarProcess.Value = Seek(29)
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
        Open App.Path & "\Extrato_Macromedidor.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Extrato Macromedidor Compesa A4/A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        Dim i2 As Integer
        Dim i3 As Integer
        Dim i4 As Integer
        For i = 1 To 2
            Reg(i).ESCRITORIO = ""
            Reg(i).INSCRICAO = ""
            Reg(i).NOME = ""
            Reg(i).MATRICULA = ""
            Reg(i).ENDERECO = ""
            Reg(i).ENDERECO2 = ""
            Reg(i).FATURA = ""
            Reg(i).RESPONSAVEL = ""
            Reg(i).ENDERECO_ENTREGA = ""
            For i2 = 1 To 3
                Reg(i).ULTIMOS_CONSUMOS(i2) = ""
            Next
            Reg(i).DATA_LEITURA = ""
            Reg(i).LEITURA = ""
            Reg(i).CONSUMO_TOTAL = ""
            Reg(i).NUMERO_UNIDADES = ""
            Reg(i).CONSUMO_UNIDADE = ""
            Reg(i).COD_AUXILIAR = ""
            For i3 = 1 To 3
                Reg(i).SITUACAO_FATURAMENTO(i3) = ""
            Next
            For i4 = 1 To 4
                Reg(i).EXTRATO(i4) = ""
            Next
            Reg(i).SEQ = ""
            Reg(i).GRUPO = ""
            Reg(i).FIRMA = ""
        Next
End Sub
