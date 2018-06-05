VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdem_Substituicao_Hid_A5"
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
'Aplicação: Ordem de Substituição de Hidrometro Compesa A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Ordem de Substituição de Hidrometro Compesa A4/A5"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        NUMERO As String
        EMISSAO As String
        VALIDADE As String
        MATRICULA As String
        INSCRICAO As String
        TELEFONE As String
        NOME As String
        CNPJ_CPF As String
        RES As String
        COM As String
        IND As String
        PUB As String
        TIPO_CLIENTE As String
        ENDERECO As String
        ESGOTO As String
        ESGOTO_FIXO As String
        LIGACAO_AGUA As String
        MARCA As String
        CAPACIDADE As String
        NUMERO_HIDROMETRO As String
        DIAMETRO As String
        LOCAL_INSTALACAO As String
        LEITURA As String
        DATA_INSTALACAO As String
        ANORMALIDADE As String
        PROTECAO As String
        CAVALETE As String
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Integer
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
                    posX = 150
                    posY = 155
                Case 2
                    posX = 1900
                    posY = 150
            End Select
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 1060 & " pontos -" & posY & " pontos moveto (" & Reg(i).NUMERO & ") show"
            posY = posY + 45
            Print #outFN, posX + 1180 & " pontos -" & posY & " pontos moveto (" & Reg(i).EMISSAO & ") show"
            posY = posY + 45
            Print #outFN, posX + 1290 & " pontos -" & posY & " pontos moveto (" & Reg(i).VALIDADE & ") show"
            posY = posY + 75
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & "                 " & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 1110 & " pontos -" & posY & " pontos moveto (" & Reg(i).TELEFONE & ") show"
            posY = posY + 65
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 290 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 1110 & " pontos -" & posY & " pontos moveto (" & Reg(i).CNPJ_CPF & ") show"
            posY = posY + 65
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).RES & ") show"
            Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg(i).COM & ") show"
            Print #outFN, posX + 690 & " pontos -" & posY & " pontos moveto (" & Reg(i).IND & ") show"
            Print #outFN, posX + 880 & " pontos -" & posY & " pontos moveto (" & Reg(i).PUB & ") show"
            Print #outFN, posX + 1190 & " pontos -" & posY & " pontos moveto (" & Reg(i).TIPO_CLIENTE & ") show"
            posY = posY + 65
            Print #outFN, posX + 190 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & ") show"
            posY = posY + 68
            Print #outFN, posX + 285 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESGOTO & ") show"
            Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg(i).ESGOTO_FIXO & ") show"
            Print #outFN, posX + 1050 & " pontos -" & posY & " pontos moveto (" & Reg(i).LIGACAO_AGUA & ") show"
            posY = posY + 200
            Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
            Print #outFN, posX + 210 & " pontos -" & posY & " pontos moveto (" & Reg(i).MARCA & ") show"
            Print #outFN, posX + 690 & " pontos -" & posY & " pontos moveto (" & Reg(i).CAPACIDADE & ") show"
            posY = posY + 53
            Print #outFN, posX + 210 & " pontos -" & posY & " pontos moveto (" & Reg(i).NUMERO_HIDROMETRO & ") show"
            Print #outFN, posX + 690 & " pontos -" & posY & " pontos moveto (" & Reg(i).DIAMETRO & ") show"
            posY = posY + 53
            Print #outFN, posX + 365 & " pontos -" & posY & " pontos moveto (" & Reg(i).LOCAL_INSTALACAO & ") show"
            Print #outFN, posX + 690 & " pontos -" & posY & " pontos moveto (" & Reg(i).LEITURA & ") show"
            posY = posY + 53
            Print #outFN, posX + 365 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_INSTALACAO & ") show"
            posY = posY + 40
            Print #outFN, posX + 290 & " pontos -" & posY & " pontos moveto (" & Reg(i).ANORMALIDADE & ") show"
            posY = posY + 53
            Print #outFN, posX + 365 & " pontos -" & posY & " pontos moveto (" & Reg(i).PROTECAO & ") show"
            posY = posY + 53
            Print #outFN, posX + 365 & " pontos -" & posY & " pontos moveto (" & Reg(i).CAVALETE & ") show"
            posY = posY + 1370
            Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
            Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
            Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Next
        Print #outFN, "showpage"
        '----------------Fim da Parte Interna------------------------------------------------------------------------------
        Print #outFN, "save EPSForm2 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

'        If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String
        Dim wIndice As Integer
        wIndice = 0
        
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

           wIndice = wIndice + 1

           Reg(wIndice).NUMERO = Trim(Mid(linha, 1, 9))
           Reg(wIndice).EMISSAO = Trim(Mid(linha, 10, 10))
           Reg(wIndice).VALIDADE = Trim(Mid(linha, 20, 10))
           Reg(wIndice).MATRICULA = Trim(Mid(linha, 30, 9))
           Reg(wIndice).INSCRICAO = Trim(Mid(linha, 39, 20))
           Reg(wIndice).TELEFONE = Trim(Mid(linha, 59, 9))
           Reg(wIndice).NOME = Trim(Mid(linha, 68, 41))
           Reg(wIndice).CNPJ_CPF = Trim(Mid(linha, 109, 18))
           Reg(wIndice).RES = Trim(Mid(linha, 127, 4))
           Reg(wIndice).COM = Trim(Mid(linha, 131, 4))
           Reg(wIndice).IND = Trim(Mid(linha, 135, 4))
           Reg(wIndice).PUB = Trim(Mid(linha, 139, 4))
           Reg(wIndice).TIPO_CLIENTE = Trim(Mid(linha, 143, 15))
           Reg(wIndice).ENDERECO = Trim(Mid(linha, 158, 55))
           Reg(wIndice).ESGOTO = Trim(Mid(linha, 213, 10))
           Reg(wIndice).ESGOTO_FIXO = Trim(Mid(linha, 223, 5))
           Reg(wIndice).LIGACAO_AGUA = Trim(Mid(linha, 228, 15))
           Reg(wIndice).MARCA = Trim(Mid(linha, 243, 15))
           Reg(wIndice).CAPACIDADE = Trim(Mid(linha, 258, 6))
           Reg(wIndice).NUMERO_HIDROMETRO = Trim(Mid(linha, 264, 10))
           Reg(wIndice).DIAMETRO = Trim(Mid(linha, 274, 5))
           Reg(wIndice).LOCAL_INSTALACAO = Trim(Mid(linha, 279, 5))
           Reg(wIndice).LEITURA = Trim(Mid(linha, 284, 7))
           Reg(wIndice).DATA_INSTALACAO = Trim(Mid(linha, 291, 10))
           Reg(wIndice).ANORMALIDADE = Trim(Mid(linha, 310, 25))
           Reg(wIndice).PROTECAO = Trim(Mid(linha, 301, 6))
           Reg(wIndice).CAVALETE = Trim(Mid(linha, 307, 3))

           If wIndice = 2 Then
              MontaExtrato
              LimpaReg
              wIndice = 0
           End If

           frmPrincipal.txQtdReg.Text = totRegs

           'If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        If wIndice > 0 Then MontaExtrato

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
        Open App.Path & "\Ordem_Substituicao_Hidrometro_Compesa_A5_Int.ps" For Input As #200
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
        Open App.Path & "\Ordem_Substituicao_Hidrometro_Compesa_A5_Ext.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Ordem de Substituição de Hidrometro Compesa A4/A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer

        For i = 1 To 2
            Reg(i).NUMERO = ""
            Reg(i).EMISSAO = ""
            Reg(i).VALIDADE = ""
            Reg(i).MATRICULA = ""
            Reg(i).INSCRICAO = ""
            Reg(i).TELEFONE = ""
            Reg(i).NOME = ""
            Reg(i).CNPJ_CPF = ""
            Reg(i).RES = ""
            Reg(i).COM = ""
            Reg(i).IND = ""
            Reg(i).PUB = ""
            Reg(i).TIPO_CLIENTE = ""
            Reg(i).ENDERECO = ""
            Reg(i).ESGOTO = ""
            Reg(i).ESGOTO_FIXO = ""
            Reg(i).LIGACAO_AGUA = ""
            Reg(i).MARCA = ""
            Reg(i).CAPACIDADE = ""
            Reg(i).NUMERO_HIDROMETRO = ""
            Reg(i).DIAMETRO = ""
            Reg(i).LOCAL_INSTALACAO = ""
            Reg(i).LEITURA = ""
            Reg(i).DATA_INSTALACAO = ""
            Reg(i).ANORMALIDADE = ""
            Reg(i).PROTECAO = ""
            Reg(i).CAVALETE = ""
        Next
End Sub
