VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAvisodeCorte_Boleto_A5"
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
'Aplicação: Aviso de Corte CAERN A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "Compesa", AppDesc = "Aviso de Corte Boleto Compesa A4/A5"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        ESCRITORIO1 As String
        CICLO As String
        INSCRICAO As String
        NOME As String
        MATRICULA As String
        ENDERECO As String
        GRUPO As String
        NR_ORDEM As String
        A As String
        E As String
        ECONOMIAS As String
        RES As String
        COM As String
        IND As String
        PUB As String
        TIPO_CONSUMIDOR As String
        DATA_EMISSAO As String
        NAO_RECEBER_APOS As String
        FATURA_EM_ATRASO(28) As String
        VENCIMENTO(28) As String
        VALOR(28) As String
        GRUPO_FIRMA As String
        GRUPO_FIRMA2 As String
        TOTAL_DEBITO As String
        ESCRITORIO As String
        SEQUENCIA As String
        LINHA_DIGITAVEL As String
        BAR_CODE As String
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private totFacs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim posY As Long
        Dim posX As Long
        Dim i2 As Integer

        wRec = wRec + 1
        '---------------------------------------------------------------------------------------
        'Parte Interna
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        For i = 1 To 2
            If Len(Trim(Reg(i).NOME)) > 0 Then totRegs = totRegs + 1
            Select Case i
                Case 1
                  posX = 430
                  posY = 1150
                Case 2
                  posX = 2180
                  posY = 1150
            End Select
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Mid(Reg(i).ESCRITORIO, 1, 20) & ") show"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO_FIRMA & ") show"
            Print #outFN, posX + 600 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            posY = posY + 40
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (Ciclo: " & Reg(i).CICLO & ") show"
            Print #outFN, posX + 600 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESCRITORIO1 & ") show"
            posY = posY + 40
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            posY = posY + 40
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Mid(Reg(i).ENDERECO, 1, 45) & ") show"
            posY = posY + 40
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Mid(Reg(i).ENDERECO, 46, 45) & ") show"
            posY = posY + 40
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQUENCIA & ") show"
            posY = posY + 60
            Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (Rec: " & Format(wRec, "00000") & ") show"
        Next
        Print #outFN, "showpage"
        Print #outFN, " "
        '---------------------------------------------------------------------------------------
        'Parte Interna
        Print #outFN, "save EPSForm2 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        For i = 1 To 2
            Select Case i
                Case 1
                  posX = 150
                  posY = 275
                Case 2
                  posX = 1920
                  posY = 275
            End Select
            'Rota
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 260 & " pontos -" & posY & " pontos moveto (Escritório: " & Reg(i).ESCRITORIO1 & ") show"
            Print #outFN, posX + 1000 & " pontos -" & posY & " pontos moveto (Ciclo: " & Reg(i).CICLO & ") show"
            posY = posY + 90
            'Inscrição, Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 545 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 1270 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            posY = posY + 90
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            'Endereço, Grupo, Numero de Ordem
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
            Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Reg(i).NR_ORDEM & ") show"
            posY = posY + 85
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            'A, E, Economia, Com, ind, Pub, Tipo Consumidor, Data Emissão, Não Receber Após
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).A & ") show"
            Print #outFN, posX + 90 & " pontos -" & posY & " pontos moveto (" & Reg(i).E & ") show"
            Print #outFN, posX + 190 & " pontos -" & posY & " pontos moveto (" & Reg(i).RES & ") show"
            Print #outFN, posX + 320 & " pontos -" & posY & " pontos moveto (" & Reg(i).COM & ") show"
            Print #outFN, posX + 450 & " pontos -" & posY & " pontos moveto (" & Reg(i).IND & ") show"
            Print #outFN, posX + 590 & " pontos -" & posY & " pontos moveto (" & Reg(i).PUB & ") show"
            Print #outFN, posX + 200 & " pontos -" & posY + 10 & " pontos moveto (" & Reg(i).ECONOMIAS & ") show"
            Print #outFN, posX + 850 & " pontos -" & posY & " pontos moveto (" & Reg(i).TIPO_CONSUMIDOR & ") show"
            Print #outFN, posX + 1050 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_EMISSAO & ") show"
            Print #outFN, posX + 1260 & " pontos -" & posY & " pontos moveto (" & Reg(i).NAO_RECEBER_APOS & ") show"
            'Faturas em Atraso
            posY = posY + 110
            Dim wAltX As Long
            Dim wContFat As Integer
            wContFat = 0
            wAltX = 0
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
            For i2 = 1 To 28
                wContFat = wContFat + 1
                Print #outFN, posX - 50 + wAltX & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).FATURA_EM_ATRASO(i2))) & Reg(i).FATURA_EM_ATRASO(i2) & ") show"
                Print #outFN, posX + 220 + wAltX & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).VENCIMENTO(i2))) & Reg(i).VENCIMENTO(i2) & ") show"
                Print #outFN, posX + 400 + wAltX & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg(i).VALOR(i2))) & Reg(i).VALOR(i2) & ") show"
                
                posY = posY + 30
                
                If wContFat = 14 Then
                   posY = 650
                   wContFat = 0
                   wAltX = 710
                End If
            Next
            'Total Débito
            Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
            posY = 1170
            Print #outFN, posX + 1060 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg(i).TOTAL_DEBITO)) & Reg(i).TOTAL_DEBITO & ") show"
            posY = posY + 480
            'Escritorio, Sequencia
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO_FIRMA2 & ") show"
            Print #outFN, posX + 1170 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQUENCIA & ") show"
'            'Linha Digitável
            posY = posY + 165
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 450 & " pontos -" & posY & " pontos moveto (" & Reg(i).LINHA_DIGITAVEL & ") show"
            'Código de Barra
            posY = posY + 168
            Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
            Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg(i).BAR_CODE & ") show"
            'Inscrição, Não receber, Matricula, Nr Ordem
            posY = posY + 65
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 555 & " pontos -" & posY & " pontos moveto (" & Reg(i).NAO_RECEBER_APOS & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            Print #outFN, posX + 1265 & " pontos -" & posY & " pontos moveto (" & Reg(i).NR_ORDEM & ") show"
            'Total Débito
            Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
            posY = posY + 165
            Print #outFN, posX + 1060 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg(i).TOTAL_DEBITO)) & Reg(i).TOTAL_DEBITO & ") show"
            posY = posY + 200
            'Escritorio, Sequencia
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO_FIRMA2 & ") show"
            Print #outFN, posX + 1170 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQUENCIA & ") show"
            posY = posY + 130
            'Sequencia e Recuperação
            posY = posY + 10
            Print #outFN, "/Helvetica-AP findfont 4.5 scalefont setfont"
            Print #outFN, posX + 235 & " pontos -" & posY & " pontos moveto (Rec: " & Format(wRec, "000000") & "   Seq: " & Reg(i).SEQUENCIA & ") show"
        
        Next
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

'       If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wIndice As Integer
        wIndice = 1
        totRegs = 0

        'Gravar Cabeçalho
        GravaHeader

        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001

        Do While Not EOF(inFN)
           DoEvents
           Line Input #inFN, linha
           
           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")

           Reg(wIndice).ESCRITORIO1 = Trim(Mid(linha, 1, 25))
           Reg(wIndice).CICLO = Trim(Mid(linha, 26, 13))
           Reg(wIndice).INSCRICAO = Trim(Mid(linha, 39, 20))
           Reg(wIndice).NOME = Trim(Mid(linha, 59, 50))
           Reg(wIndice).MATRICULA = Trim(Mid(linha, 109, 10))
           Reg(wIndice).ENDERECO = Trim(Mid(linha, 119, 72))
           Reg(wIndice).GRUPO = Trim(Mid(linha, 191, 2))
           Reg(wIndice).NR_ORDEM = Trim(Mid(linha, 193, 9))
           Reg(wIndice).A = Trim(Mid(linha, 202, 1))
           Reg(wIndice).E = Trim(Mid(linha, 203, 1))
           Reg(wIndice).RES = Trim(Mid(linha, 204, 3))
           Reg(wIndice).COM = Trim(Mid(linha, 207, 3))
           Reg(wIndice).IND = Trim(Mid(linha, 210, 3))
           Reg(wIndice).PUB = Trim(Mid(linha, 213, 3))
           Reg(wIndice).ECONOMIAS = Trim(Mid(linha, 22, 4))
           Reg(wIndice).TIPO_CONSUMIDOR = Trim(Mid(linha, 216, 8))
           Reg(wIndice).DATA_EMISSAO = Trim(Mid(linha, 224, 10))
           Reg(wIndice).NAO_RECEBER_APOS = Trim(Mid(linha, 234, 10))
           Dim winit As Integer
           winit = 244
           For i = 1 To 28
               Reg(wIndice).FATURA_EM_ATRASO(i) = Trim(Mid(linha, winit, 11))
               winit = winit + 11
               Reg(wIndice).VENCIMENTO(i) = Trim(Mid(linha, winit, 10))
               winit = winit + 10
               Reg(wIndice).VALOR(i) = Trim(Mid(linha, winit, 14))
               winit = winit + 14
           Next
           Reg(wIndice).TOTAL_DEBITO = Trim(Mid(linha, 1224, 14))
           Reg(wIndice).GRUPO_FIRMA = Trim(Mid(linha, 1238, 19))
           Reg(wIndice).GRUPO_FIRMA2 = Trim(Mid(linha, 1238, 33))
           Reg(wIndice).SEQUENCIA = Trim(Mid(linha, 1269, 9))
           Reg(wIndice).LINHA_DIGITAVEL = Trim(Mid(linha, 1278, 55))
           Reg(wIndice).BAR_CODE = Trim(Mid(linha, 1333, 112))
           
           wIndice = wIndice + 1
           
           If wIndice = 3 Then
              wIndice = 1
              MontaExtrato
              LimpaReg
           End If

           'If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(inFN)
        Loop
        
        If wIndice > 1 Then
           MontaExtrato
           LimpaReg
        End If
        
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
       Open iFile For Input As #inFN

       pastaProcessar = pProc
       pastaGerar = pGer

       jobName = nomeDoJob(iFile)

       outFN = FreeFile
       oFile = pastaGerar & Left$(jobName, Len(jobName) - 4) & ".ps"

       jobName = nomeDoJob(oFile)

       Open oFile For Output As outFN

       totRegs = 0
       totFacs = 0

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
        Open App.Path & "\AvisoDeCorte_Boleto_Compesa_Ext.ps" For Input As #200
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
        Print #outFN, "/DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
        Print #outFN, ">> /ReusableStreamDecode filter"
        Open App.Path & "\AvisoDeCorte_Boleto_Compesa_Int.ps" For Input As #200
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
        Print #outFN, "%Aplicação: AVISO DE CORTE BOLETO COMPESA A4/A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        Dim i2 As Integer
        For i = 1 To 2
            Reg(i).ESCRITORIO1 = ""
            Reg(i).CICLO = ""
            Reg(i).INSCRICAO = ""
            Reg(i).NOME = ""
            Reg(i).MATRICULA = ""
            Reg(i).ENDERECO = ""
            Reg(i).GRUPO = ""
            Reg(i).NR_ORDEM = ""
            Reg(i).A = ""
            Reg(i).E = ""
            Reg(i).ECONOMIAS = ""
            Reg(i).RES = ""
            Reg(i).COM = ""
            Reg(i).IND = ""
            Reg(i).PUB = ""
            Reg(i).TIPO_CONSUMIDOR = ""
            Reg(i).DATA_EMISSAO = ""
            Reg(i).NAO_RECEBER_APOS = ""
            For i2 = 1 To 28
                Reg(i).FATURA_EM_ATRASO(i2) = ""
                Reg(i).VENCIMENTO(i2) = ""
                Reg(i).VALOR(i2) = ""
            Next
            Reg(i).GRUPO_FIRMA = ""
            Reg(i).TOTAL_DEBITO = ""
            Reg(i).ESCRITORIO = ""
            Reg(i).SEQUENCIA = ""
            Reg(i).LINHA_DIGITAVEL = ""
            Reg(i).BAR_CODE = ""
        Next
End Sub
