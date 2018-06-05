VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAvisodeCob_e_Perda_Benef_A4"
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
'Aplicação: Aviso de Cobrança e Perda de Beneficio Compesa A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Aviso de Cobrança e Perda de Beneficio Compesa A4"

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
        VENCIMENTO_BOLETO As String
        BAR_CODE As String
        FATURA(12) As String
        VENCIMENTO(12) As String
        VALOR(12) As String
        VALOR_FICHA As String
        VALOR_TOTAL As String
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
        posY = 375
        
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
        
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
        posY = posY + 510
        For i = 1 To 12
            Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FATURA(i))) & Reg.FATURA(i) & ") show"
            Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO(i) & ") show"
            Print #outFN, posX + 1050 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR(i))) & Reg.VALOR(i) & ") show"

            posY = posY + 31
        Next
        posY = 1570
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_TOTAL)) & Reg.VALOR_TOTAL & ") show"
        posY = posY + 450
        'Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        'Print #outFN, posX + 1270 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO_BOLETO & ") show"

        posY = posY + 1030
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg.LINHA_DIGITAVEL & ") show"
        posY = posY + 180
        Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.BAR_CODE & ") show"
        posY = posY + 95
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.ORDEM & ") show"
        Print #outFN, posX + 1650 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 120
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 590 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO_BOLETO & ") show"
        Print #outFN, posX + 1480 & " pontos -" & posY & " pontos moveto (" & Reg.VALOR_FICHA & ") show"
        Print #outFN, "showpage"
        'Fim da Parte Interna
        '**********Inicio Parte Externa************************************************
        posX = 810
        posY = 600
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
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        
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

           For i1 = 1 To 15
               If i1 = 1 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 502, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 513, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 523, 14))
               End If
               If i1 = 2 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 537, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 548, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 558, 14))
               End If
               If i1 = 3 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 572, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 583, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 593, 14))
               End If
               If i1 = 4 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 607, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 618, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 628, 14))
               End If
               If i1 = 5 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 642, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 653, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 663, 14))
               End If
               If i1 = 6 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 677, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 688, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 698, 14))
               End If
               If i1 = 7 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 712, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 723, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 733, 14))
               End If
               If i1 = 8 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 747, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 758, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 768, 14))
               End If
               If i1 = 9 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 782, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 793, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 803, 14))
               End If
               If i1 = 10 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 817, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 828, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 838, 14))
               End If
               If i1 = 11 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 852, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 863, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 873, 14))
               End If
               If i1 = 12 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 887, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 898, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 908, 14))
               End If
           Next
           Reg.VALOR_TOTAL = Trim(Mid(linha, 992, 14))
           Reg.VALOR_FICHA = Trim(Mid(linha, 1006, 14))

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
        Open App.Path & "\Aviso_De_Cobranca_E_Perda_de_Beneficio_Compesa_A4_Int.ps" For Input As #200
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
        Open App.Path & "\Aviso_De_Cobranca_E_Perda_de_Beneficio_Compesa_A4_Ext.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Aviso de Cobrança e Perda de Beneficio Compesa A4"
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
        Reg.VALOR_FICHA = ""
        Reg.VALOR_TOTAL = ""
        For i = 1 To 12
            Reg.FATURA(i) = ""
            Reg.VENCIMENTO(i) = ""
            Reg.VALOR(i) = ""
        Next
End Sub
