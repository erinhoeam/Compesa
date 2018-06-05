VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInspecaoDeEsgoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap Consultoria
'                 Analista de Sistemas
'Data: 08 de Junho de 2009
'Contatos - Email: erinhoeam@hotmail.com
'                  aziel@sistemap.com.br
'           Fone: (81) 9832.6084
'Aplicação: Formulário de Recadastramento Ligação Compesa A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "Compesa", AppDesc = "Formulário de Inspeção de Esgoto"

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
        MATRICULA As String
        ENDERECO As String
        BAIRRO As String
        MUNICIPIO As String
        UF As String
        CEP As String
        N_OS As String
        RES As String
        COM As String
        IND As String
        PUB As String
        TOTAL_CATEGORIAS As String
        DATA_EMISSAO As String
        PERFIL_CLIENTE As String
        ULTIMA_ALTERACAO As String
        GRUPO_COBRANCA As String
        SEQUENCIAL As String
        SITUACAO_AGUA As String
        CONSUMO_MEDIO As String
        SITUACAO_ESGOTO As String
        CONSUMO_FIXO As String
        DATA_CORTE As String
        DATA_SUPRESSAO_TOTAL As String
        ORIGEM As String
        OCORRENCIA As String
        DATA_POSICAO_DEBITO As String
        VALOR_TOTAL_DEBITO As String
        EXECUTOR As String
        CAS As String
        CPF_CNPJ As String
        RG As String
        UF2 As String
        DDD As String
        N_FONTE As String
        FONE As String
        RAMAL As String
        TIPO As String
        SERVICOS_ATUALIZACOES As String
End Type
'Variáveis gerais para os tipos
Private Reg As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim posY As Long
        Dim posX As Long

        posX = 10
        posY = 110

        If Len(Trim(Reg.NOME)) > 0 Then wRec = wRec + 1
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE

       'Executor, Escritorio
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.N_OS & ") show"
        posY = posY + 90
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO & ") show"
        posY = posY + 70
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 80
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.ORIGEM & ") show"
        posY = posY - 45
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.EXECUTOR & ") show"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg.CAS & ") show"
        posY = posY + 185
        'Inscrição, Mátricula
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 850 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (" & Reg.PERFIL_CLIENTE & ") show"
        posY = posY + 80
        Print #outFN, posX + 270 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        Print #outFN, posX + 270 & " pontos -" & posY + 35 & " pontos moveto (" & Reg.BAIRRO & " " & Reg.MUNICIPIO & "/" & Reg.UF & " CEP: " & Reg.CEP & ") show"
        posY = posY + 120
        Print #outFN, posX + 260 & " pontos -" & posY & " pontos moveto (" & Reg.RES & ") show"
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg.COM & ") show"
        Print #outFN, posX + 540 & " pontos -" & posY & " pontos moveto (" & Reg.IND & ") show"
        Print #outFN, posX + 680 & " pontos -" & posY & " pontos moveto (" & Reg.PUB & ") show"
        Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Reg.TOTAL_CATEGORIAS & ") show"
        Print #outFN, posX + 990 & " pontos -" & posY & " pontos moveto (" & Reg.ULTIMA_ALTERACAO & ") show"
        Print #outFN, posX + 1430 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO_COBRANCA & ") show"
        posY = posY + 80
        Print #outFN, posX + 460 & " pontos -" & posY & " pontos moveto (" & Reg.SITUACAO_AGUA & ") show"
        Print #outFN, posX + 1140 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_MEDIO & ") show"
        posY = posY + 80
        Print #outFN, posX + 240 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_CORTE & ") show"
        Print #outFN, posX + 1140 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_SUPRESSAO_TOTAL & ") show"
        posY = posY + 80
        Print #outFN, posX + 460 & " pontos -" & posY & " pontos moveto (" & Reg.SITUACAO_ESGOTO & ") show"
        Print #outFN, posX + 1140 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_FIXO & ") show"
        posY = posY + 45
        Print #outFN, posX + 320 & " pontos -" & posY & " pontos moveto (" & Reg.OCORRENCIA & ") show"
        posY = posY + 80
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1290 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.SERVICOS_ATUALIZACOES)) & Reg.SERVICOS_ATUALIZACOES & ") show"
        posY = posY + 80
        Print #outFN, posX + 1290 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.VALOR_TOTAL_DEBITO)) & Reg.VALOR_TOTAL_DEBITO & ") show"
        Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_POSICAO_DEBITO & ") show"
        posY = posY + 170
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 120
        Print #outFN, posX + 250 & " pontos -" & posY & " pontos moveto (" & Reg.CPF_CNPJ & ") show"
        Print #outFN, posX + 1030 & " pontos -" & posY & " pontos moveto (" & Reg.RG & ") show"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Reg.UF2 & ") show"
        posY = posY + 90
        Print #outFN, posX + 350 & " pontos -" & posY & " pontos moveto (" & Reg.DDD & ") show"
        Print #outFN, posX + 650 & " pontos -" & posY & " pontos moveto (" & Reg.FONE & ") show"
        Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (" & Reg.RAMAL & ") show"
        Print #outFN, posX + 1577 & " pontos -" & posY & " pontos moveto (" & Reg.TIPO & ") show"
        posY = posY + 930
        'Sequencia e Recuperação
        Print #outFN, "/Helvetica-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX + 100 & " pontos -" & posY & " pontos moveto (Rec: " & Format(wRec, "000000") & "   Seq: " & Reg.SEQUENCIAL & ") show"
        '**************OUTRA LADO********************
        posY = posY - 2120
        posX = posX + 1600
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.N_OS & ") show"
        posY = posY + 90
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_EMISSAO & ") show"
        posY = posY + 70
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 90
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.ORIGEM & ") show"
        posY = posY - 45
        Print #outFN, posX + 450 & " pontos -" & posY & " pontos moveto (" & Reg.EXECUTOR & ") show"
        Print #outFN, posX + 750 & " pontos -" & posY & " pontos moveto (" & Reg.CAS & ") show"
        posX = posX + 160
        posY = posY + 225
        'Inscrição, Mátricula
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 850 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (" & Reg.PERFIL_CLIENTE & ") show"
        posY = posY + 80
        Print #outFN, posX + 270 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        Print #outFN, posX + 270 & " pontos -" & posY + 35 & " pontos moveto (" & Reg.BAIRRO & " " & Reg.MUNICIPIO & "/" & Reg.UF & " CEP: " & Reg.CEP & ") show"
        posY = posY + 120
        Print #outFN, posX + 260 & " pontos -" & posY & " pontos moveto (" & Reg.RES & ") show"
        Print #outFN, posX + 410 & " pontos -" & posY & " pontos moveto (" & Reg.COM & ") show"
        Print #outFN, posX + 560 & " pontos -" & posY & " pontos moveto (" & Reg.IND & ") show"
        Print #outFN, posX + 710 & " pontos -" & posY & " pontos moveto (" & Reg.PUB & ") show"
        Print #outFN, posX + 850 & " pontos -" & posY & " pontos moveto (" & Reg.TOTAL_CATEGORIAS & ") show"
        Print #outFN, posX + 1100 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO_COBRANCA & ") show"
        posY = posY + 200
        'Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
        'Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.SERVICOS_ATUALIZACOES)) & Reg.SERVICOS_ATUALIZACOES & ") show"
        'posY = posY + 80
        'Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Space(20 - Len(Reg.VALOR_TOTAL_DEBITO)) & Reg.VALOR_TOTAL_DEBITO & ") show"
        'Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_POSICAO_DEBITO & ") show"
        'posY = posY + 280
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 380 & " pontos -" & posY - 45 & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 1290
        'Sequencia e Recuperação
        Print #outFN, "/Helvetica-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX + 100 & " pontos -" & posY & " pontos moveto (Rec: " & Format(wRec, "000000") & "   Seq: " & Reg.SEQUENCIAL & ") show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "
        'If wRec = 1 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String
        totRegs = 0

        'Gravar Cabeçalho
        GravaHeader

        Do While Not EOF(inFN)
           DoEvents
           Line Input #inFN, linha
           
           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")

           Reg.INSCRICAO = Trim(Mid(linha, 1, 20))
           Reg.NOME = Trim(Mid(linha, 21, 50))
           Reg.MATRICULA = Trim(Mid(linha, 71, 9))
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
           Reg.TOTAL_CATEGORIAS = Trim(Mid(linha, 272, 3))
           Reg.DATA_EMISSAO = Trim(Mid(linha, 275, 10))
           Reg.PERFIL_CLIENTE = Trim(Mid(linha, 285, 20))
           Reg.ULTIMA_ALTERACAO = Trim(Mid(linha, 305, 10))
           Reg.GRUPO_COBRANCA = Trim(Mid(linha, 315, 2))
           Reg.SEQUENCIAL = Trim(Mid(linha, 317, 7))
           Reg.SITUACAO_AGUA = Trim(Mid(linha, 324, 20))
           Reg.CONSUMO_MEDIO = Trim(Mid(linha, 344, 28))
           Reg.SITUACAO_ESGOTO = Trim(Mid(linha, 372, 20))
           Reg.CONSUMO_FIXO = Trim(Mid(linha, 392, 10))
           Reg.DATA_CORTE = Trim(Mid(linha, 402, 10))
           Reg.DATA_SUPRESSAO_TOTAL = Trim(Mid(linha, 412, 10))
           Reg.ORIGEM = Trim(Mid(linha, 422, 15))
           Reg.OCORRENCIA = Trim(Mid(linha, 437, 33))
           Reg.DATA_POSICAO_DEBITO = Trim(Mid(linha, 471, 10))
           Reg.VALOR_TOTAL_DEBITO = Trim(Mid(linha, 481, 18))
           Reg.EXECUTOR = Trim(Mid(linha, 499, 10))
           Reg.CAS = Trim(Mid(linha, 509, 50))
           Reg.CPF_CNPJ = Trim(Mid(linha, 559, 18))
           Reg.RG = Trim(Mid(linha, 577, 18))
           Reg.UF2 = Trim(Mid(linha, 597, 2))
           Reg.DDD = Trim(Mid(linha, 599, 3))
           Reg.FONE = Trim(Mid(linha, 602, 8))
           Reg.RAMAL = Trim(Mid(linha, 610, 4))
           Reg.TIPO = Trim(Mid(linha, 614, 3))
           Reg.SERVICOS_ATUALIZACOES = Trim(Mid(linha, 617, 17))
           
           If Len(Trim(Reg.NOME)) > 0 Then
              totRegs = totRegs + 1
              MontaExtrato
              LimpaReg
              frmPrincipal.txQtdReg.Text = totRegs
           End If

           'If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
           '   frmPrincipal.ProgressBarProcess.Value = Seek(inFN)
        Loop
        Screen.MousePointer = vbDefault

        frmResultados.txtTotCli_STL.Text = totRegs
        frmResultados.txtTotReg_STL.Text = wRec
        frmResultados.txArquivo.Text = wArq

        frmResultados.txtAplicacao = AppDesc

        Close #outFN
        Close #outFN

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
        Open App.Path & "\Inspecao_de_Esgoto.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Fomrmulário de Inspeção Compesa A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        With Reg
            .INSCRICAO = ""
            .NOME = ""
            .MATRICULA = ""
            .ENDERECO = ""
            .BAIRRO = ""
            .MUNICIPIO = ""
            .UF = ""
            .CEP = ""
            .N_OS = ""
            .RES = ""
            .COM = ""
            .IND = ""
            .PUB = ""
            .TOTAL_CATEGORIAS = ""
            .DATA_EMISSAO = ""
            .PERFIL_CLIENTE = ""
            .ULTIMA_ALTERACAO = ""
            .GRUPO_COBRANCA = ""
            .SEQUENCIAL = ""
            .SITUACAO_AGUA = ""
            .CONSUMO_MEDIO = ""
            .SITUACAO_ESGOTO = ""
            .CONSUMO_FIXO = ""
            .DATA_CORTE = ""
            .DATA_SUPRESSAO_TOTAL = ""
            .ORIGEM = ""
            .OCORRENCIA = ""
            .DATA_POSICAO_DEBITO = ""
            .VALOR_TOTAL_DEBITO = ""
            .EXECUTOR = ""
            .CAS = ""
            .CPF_CNPJ = ""
            .RG = ""
            .UF2 = ""
            .DDD = ""
            .N_FONTE = ""
            .FONE = ""
            .RAMAL = ""
            .TIPO = ""
            .SERVICOS_ATUALIZACOES = ""
        End With
End Sub




