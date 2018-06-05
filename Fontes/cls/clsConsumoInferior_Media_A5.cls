VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsumoInferior_Media_A5"
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
'Aplicação: Consumo Inferior a Media Compesa A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Consumo Inferior a Media Compesa A4/A5"

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
        MATRICULA As String
        CLIENTE As String
        INSCRICAO As String
        ENDERECO As String
        FATURA As String
        RESPONSAVEL As String
        ENDERECO_ENTREGA As String
        ULTIMOS_CONSUMOS(2) As String
        DATA_LEITURA As String
        LEITURA_ATUAL As String
        CONSUMO_ATUAL As String
        MEDIA_CONSUMO As String
        SEQ As String
        GRUPO As String
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim i2 As Long
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
                  posX = 150
                  posY = 430
                Case 2
                  posX = 1910
                  posY = 430
            End Select
            'Rota
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).FATURA & ") show"
            Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            posY = posY + 60
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            posY = posY + 75
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).CLIENTE & ") show"
            posY = posY + 65
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
            posY = posY + 85
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).RESPONSAVEL & ") show"
            posY = posY + 75
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO_ENTREGA & ") show"
            posY = posY + 320
            Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_LEITURA & ") show"
            Print #outFN, posX + 1050 & " pontos -" & posY & " pontos moveto (" & Reg(i).LEITURA_ATUAL & ") show"
            posY = posY + 40
            Print #outFN, posX + 680 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO_ATUAL & ") show"
            posY = posY + 90
            Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Reg(i).MEDIA_CONSUMO & ") show"
            posY = posY + 245
            Print #outFN, posX + 190 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESCRITORIO & ") show"
            posY = posY + 180
            Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8 scalefont setfont"
            For i2 = 1 To 2
                Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).ULTIMOS_CONSUMOS(i2) & ") show"
                posY = posY + 30
            Next
            posY = posY + 180
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQ & ") show"
            posY = posY + 35
            Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
            Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
            posY = posY + 500
            Print #outFN, posX + 100 & " pontos -" & posY & " pontos moveto (Rec: " & Format(wRec, "00000") & ") show"
            Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (Arq: " & wArq & ") show"
        Next
        Print #outFN, "showpage"
        'Fim da Parte Interna
        '**********Inicio Parte Externa************************************************
        Print #outFN, "save EPSForm2 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

'       If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wCont As Integer, wIndice As Integer
        Dim wFlagDetalhe As Boolean
        wRec = 0
        wCont = 0
        wIndice = 1
        totRegs = 0

        'Dividir Registros
        Open pastaProcessar & "temporario1.tmp" For Output As #24
        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha
           
           If Len(Trim(Mid(linha, 2, 90))) = 0 And Mid(linha, 1, 1) = "1" Then
              Beep
           Else
              Print #24, linha
           End If

           'If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
           '  frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        Close #24
        Close #inFN
        
        'Dividir Registros
        Open pastaProcessar & "temporario1.tmp" For Input As #25
        Open pastaProcessar & "temporario.tmp" For Output As #26
        'frmPrincipal.ProgressBarProcess.Max = LOF(25)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(25)
           DoEvents
           Line Input #25, linha
           wCont = wCont + 1

           If wCont = 14 Then
              Print #26, linha
              Print #26, ""
              wCont = 0
           Else
              Print #26, linha
           End If

           'If Seek(25) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(25)

        Loop
        Close #25
        Close #26

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
                  Case 1
                       Reg(wIndice).FATURA = Trim(Mid(linha, 2, 9))
                       Reg(wIndice).MATRICULA = Trim(Mid(linha, 37, 10))
                  Case 2
                       Reg(wIndice).INSCRICAO = Trim(Mid(linha, 2, 20))
                  Case 3
                       Reg(wIndice).CLIENTE = Trim(Mid(linha, 2, 60))
                  Case 4
                       Reg(wIndice).ENDERECO = Trim(Mid(linha, 2, 60))
                  Case 5
                       Reg(wIndice).RESPONSAVEL = Trim(Mid(linha, 2, 60))
                  Case 6
                       Reg(wIndice).ENDERECO_ENTREGA = Trim(Mid(linha, 2, 60))
                  Case 7
                       Reg(wIndice).DATA_LEITURA = Trim(Mid(linha, 7, 10))
                       Reg(wIndice).LEITURA_ATUAL = Trim(Mid(linha, 42, 14))
                  Case 8
                       Reg(wIndice).CONSUMO_ATUAL = Trim(Mid(linha, 14, 14))
                  Case 9
                       Reg(wIndice).MEDIA_CONSUMO = Trim(Mid(linha, 30, 14))
                  Case 10
                       Reg(wIndice).ESCRITORIO = Trim(Mid(linha, 2, 60))
                  Case 11
                       Reg(wIndice).ULTIMOS_CONSUMOS(1) = Trim(Mid(linha, 2, 60))
                  Case 12
                       Reg(wIndice).ULTIMOS_CONSUMOS(2) = Trim(Mid(linha, 2, 60))
                  Case 13
                       Reg(wIndice).SEQ = Trim(Mid(linha, 2, 60))
                  Case 14
                       Reg(wIndice).GRUPO = Trim(Mid(linha, 2, 60))
           End Select

           If Mid(linha, 1, 1) = "" Then
              wIndice = wIndice + 1
              wCont = 0
           End If
           
           If wIndice = 3 Then
              wIndice = 1
              wCont = 0
              MontaExtrato
              frmPrincipal.txQtdReg.Text = totRegs
              LimpaReg
           End If

           'If Seek(29) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(29)
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
        Open App.Path & "\Consumo_Abaixo_da_Media_Int.ps" For Input As #200
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
        Open App.Path & "\Consumo_Abaixo_da_Media_Ext.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Consumo Inferior a Media Compesa A4/A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        For i = 1 To 2
            Reg(i).MATRICULA = ""
            Reg(i).CLIENTE = ""
            Reg(i).INSCRICAO = ""
            Reg(i).ENDERECO = ""
            Reg(i).FATURA = ""
            Reg(i).RESPONSAVEL = ""
            Reg(i).ENDERECO_ENTREGA = ""
            Reg(i).ULTIMOS_CONSUMOS(1) = ""
            Reg(i).ULTIMOS_CONSUMOS(2) = ""
            Reg(i).ESCRITORIO = ""
            Reg(i).LEITURA_ATUAL = ""
            Reg(i).CONSUMO_ATUAL = ""
            Reg(i).MEDIA_CONSUMO = ""
            Reg(i).SEQ = ""
            Reg(i).GRUPO = ""
        Next
End Sub
