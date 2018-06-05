VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdem_Corte_New_Compesa_A4"
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
'Aplicação: Ordem de Corte Compesa A4 Novo Form
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Ordem de Corte Compesa A4 Novo Form"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        INSCRICAO As String
        NOME As String
        ENDERECO As String
        BAIRRO As String
        MUNICIPIO As String
        UF As String
        CEP As String
        N_OS As String
        MATRICULA As String
        RES  As String
        COM As String
        IND As String
        PUB As String
        A As String
        E As String
        DATA_EMISSAO_ORDEM As String
        GRUPO_COBRANCA As String
        N_HIDROMETRO As String
        DATA_EMISSAO_AVISO As String
        SEQUENCIAL As String
        CICLO As String
        PERFIL_CLIENTE As String
        CAS As String
        QTD_CONTAS_DEBITO As String
        FATURA(12) As String
        VENCIMENTO(12) As String
        VALOR(12) As String
        VALOR_TOTAL As String
        LINHA_DIGITAVEL As String
        CODE_BAR As String
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
        posX = 210
        posY = 440
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        'Rota
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        posY = posY + 87
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Trim(Reg.ENDERECO) & ") show"
        posY = posY + 30
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & " " & Reg.MUNICIPIO & "/" & Reg.UF & "  " & Reg.CEP & ") show"

        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.N_OS & ") show"
        posY = posY + 90
        Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg.RES & ") show"
        Print #outFN, posX + 130 & " pontos -" & posY & " pontos moveto (" & Reg.COM & ") show"
        Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Reg.IND & ") show"
        Print #outFN, posX + 330 & " pontos -" & posY & " pontos moveto (" & Reg.PUB & ") show"
        Print #outFN, posX + 480 & " pontos -" & posY & " pontos moveto (" & Reg.A & ") show"
        Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg.E & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO_ORDEM & ") show"
        Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg.N_HIDROMETRO & ") show"
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO_COBRANCA & ") show"
        Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO_AVISO & ") show"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 550
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.CICLO & ") show"
        Print #outFN, posX + 430 & " pontos -" & posY & " pontos moveto (" & Reg.CAS & ") show"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.PERFIL_CLIENTE & ") show"
        posY = posY + 110
        '**********ORDEM DE CORTE**********************************************************************************
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        posY = posY + 87
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Trim(Reg.ENDERECO) & ") show"
        posY = posY + 30
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & " " & Reg.MUNICIPIO & "/" & Reg.UF & "  " & Reg.CEP & ") show"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.N_OS & ") show"
        posY = posY + 90
        Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg.RES & ") show"
        Print #outFN, posX + 130 & " pontos -" & posY & " pontos moveto (" & Reg.COM & ") show"
        Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Reg.IND & ") show"
        Print #outFN, posX + 330 & " pontos -" & posY & " pontos moveto (" & Reg.PUB & ") show"
        Print #outFN, posX + 480 & " pontos -" & posY & " pontos moveto (" & Reg.A & ") show"
        Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg.E & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO_ORDEM & ") show"
        Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg.N_HIDROMETRO & ") show"
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO_COBRANCA & ") show"
        Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO_AVISO & ") show"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 110
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
        For i = 1 To 12

            Print #outFN, posX - 100 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FATURA(i))) & Reg.FATURA(i) & ") show"
            Print #outFN, posX + 220 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO(i) & ") show"
            Print #outFN, posX + 340 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR(i))) & Reg.VALOR(i) & ") show"

            posY = posY + 31
        Next
        posY = 2020
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 340 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_TOTAL)) & Reg.VALOR_TOTAL & ") show"
        posY = posY + 370
        '********************ORDEM DE CORTE VIA CLIENTE**************************************************************
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        posY = posY + 77
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Trim(Reg.ENDERECO) & ") show"
        posY = posY + 30
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & " " & Reg.MUNICIPIO & "/" & Reg.UF & "  " & Reg.CEP & ") show"
        Print #outFN, posX + 1630 & " pontos -" & posY & " pontos moveto (" & Reg.N_OS & ") show"
        posY = posY + 90
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
        For i = 1 To 12

            Print #outFN, posX - 100 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FATURA(i))) & Reg.FATURA(i) & ") show"
            Print #outFN, posX + 570 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO(i) & ") show"
            Print #outFN, posX + 890 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR(i))) & Reg.VALOR(i) & ") show"

            posY = posY + 29
        Next
        posY = 2910
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 1520 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_TOTAL)) & Reg.VALOR_TOTAL & ") show"
        posY = posY + 225
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg.LINHA_DIGITAVEL & ") show"
        posY = posY + 180
        Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.CODE_BAR & ") show"
        posY = posY + 180
        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & "   Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

        'If wRec = 20 Then End
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
           
           Reg.INSCRICAO = Trim(Mid(linha, 1, 20))
           Reg.NOME = Trim(Mid(linha, 21, 50))
           Reg.MATRICULA = Trim(Mid(linha, 71, 8)) & "." & Trim(Mid(linha, 79, 1))
           Reg.ENDERECO = Trim(Mid(linha, 80, 100))
           Reg.BAIRRO = Trim(Mid(linha, 180, 30))
           Reg.MUNICIPIO = Trim(Mid(linha, 210, 30))
           Reg.UF = Trim(Mid(linha, 240, 2))
           Reg.CEP = Trim(Mid(linha, 242, 9))
           Reg.N_OS = Trim(Mid(linha, 251, 9))
           Reg.RES = Trim(Mid(linha, 260, 3))
           Reg.COM = Trim(Mid(linha, 263, 3))
           Reg.IND = Trim(Mid(linha, 266, 3))
           Reg.PUB = Trim(Mid(linha, 269, 3))
           Reg.A = Trim(Mid(linha, 272, 1))
           Reg.E = Trim(Mid(linha, 273, 1))
           Reg.DATA_EMISSAO_ORDEM = Trim(Mid(linha, 274, 10))
           Reg.N_HIDROMETRO = Trim(Mid(linha, 284, 10))
           Reg.GRUPO_COBRANCA = Trim(Mid(linha, 294, 2))
           Reg.DATA_EMISSAO_AVISO = Trim(Mid(linha, 296, 10))
           Reg.SEQUENCIAL = Trim(Mid(linha, 306, 7))
           Reg.CICLO = Trim(Mid(linha, 313, 7))
           Reg.PERFIL_CLIENTE = Trim(Mid(linha, 320, 20))
           Reg.CAS = Trim(Mid(linha, 340, 50))
           Reg.QTD_CONTAS_DEBITO = Trim(Mid(linha, 390, 3))

           For i1 = 1 To 12
               If i1 = 1 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 393, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 400, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 410, 14))
               End If
               If i1 = 2 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 424, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 431, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 441, 14))
               End If
               If i1 = 3 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 455, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 462, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 472, 14))
               End If
               If i1 = 4 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 486, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 493, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 503, 14))
               End If
               If i1 = 5 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 517, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 524, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 534, 14))
               End If
               If i1 = 6 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 548, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 555, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 565, 14))
               End If
               If i1 = 7 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 579, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 586, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 596, 14))
               End If
               If i1 = 8 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 610, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 617, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 627, 14))
               End If
               If i1 = 9 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 641, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 648, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 658, 14))
               End If
               If i1 = 10 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 672, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 679, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 689, 14))
               End If
               If i1 = 11 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 703, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 710, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 720, 14))
               End If
               If i1 = 12 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 734, 7))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 741, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 751, 14))
               End If
           Next
           Reg.VALOR_TOTAL = Trim(Mid(linha, 765, 14))
           Reg.LINHA_DIGITAVEL = Trim(Mid(linha, 779, 55))
           Reg.CODE_BAR = Trim(Mid(linha, 834, 112))

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
        Open App.Path & "\COMUNICADO_DE_CORTE_COMPESA_A4.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Comunicado de Corte A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        With Reg
            .INSCRICAO = ""
            .NOME = ""
            .ENDERECO = ""
            .BAIRRO = ""
            .MUNICIPIO = ""
            .UF = ""
            .CEP = ""
            .N_OS = ""
            .MATRICULA = ""
            .RES = ""
            .COM = ""
            .IND = ""
            .PUB = ""
            .A = ""
            .E = ""
            .DATA_EMISSAO_ORDEM = ""
            .GRUPO_COBRANCA = ""
            .N_HIDROMETRO = ""
            .DATA_EMISSAO_AVISO = ""
            .SEQUENCIAL = ""
            .CICLO = ""
            .PERFIL_CLIENTE = ""
            .CAS = ""
            .QTD_CONTAS_DEBITO = ""
            For i = 1 To 12
                .FATURA(i) = ""
                .VENCIMENTO(i) = ""
                .VALOR(i) = ""
            Next
            .VALOR_TOTAL = ""
            .LINHA_DIGITAVEL = ""
            .CODE_BAR = ""
        End With
End Sub
