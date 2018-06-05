VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoletim_Cadastral_Compesa_A4"
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
'Aplicação: Boletim Cadastral Compesa A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Boletim Cadastral Compesa A4"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        'Dados Cliente
        COD_CLIENTE As String
        NOME As String
        TIPO_CLIENTE As String
        CNPJ_CPF As String
        RG As String
        DATA_EMISSAO_RG As String
        ORG_EXPEDITOR As String
        UF_ORG_EXPEDIDOR As String
        PROFISSAO As String
        SEXO As String
        NOME_MAE As String
        ATIVO_INATIVO As String
        DATA_NASCIMENTO As String
        TIPO_ENDERECO As String
        COD_LOGRADOURO As String
        ENDERECO As String
        COMPLEMENTO As String
        CEP As String
        BAIRRO As String
        REF As String
        NUMERO As String
        TIPO_TELEFONE1 As String
        DDD1 As String
        TELEFONE1 As String
        RAMAL1 As String
        TIPO_TELEFONE2 As String
        DDD2 As String
        TELEFONE2 As String
        RAMAL2 As String
        'Dados Imóvel
        INSCRICAO As String
        MATRICULA As String
        N_MORADORES As String
        NOME_CONTA As String
        COD_CLIENTE_PROPRIETARIO As String
        COD_CLIENTE_USUARIO As String
        COD_LOGRADOURO_IMOVEL As String
        ENDERECO_IMOVEL As String
        COMPLEMENTO_IMOVEL As String
        CEP_IMOVEL As String
        BAIRRO_IMOVEL As String
        REF_IMOVEL As String
        NUMERO_IMOVEL As String
        CATEGORIA As String
        QTD_ECONOMIA As String
        AGUA As String
        DIAM As String
        MAT As String
        INF As String
        SUP As String
        PISCINA As String
        JARDIM As String
        PAVIMENTECAO As String
        RUA As String
        POCO As String
        FONTE_ABASTECIMENTO As String
        N_PONTOS As String
        SIT_ESGOTO As String
        DIA_ESGOTO As String
        MAT_ESGOTO As String
        PERFIL_IMOVEL As String
        DESPEJO As String
        LEITURA_INICIAL As String
        CAP As String
        MARCA As String
        LOCAL As String
        PROTECAO As String
        CAVALETE As String
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
        posY = 1115
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        'Rota
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 130 & " pontos -" & posY & " pontos moveto (" & Reg.COD_CLIENTE & ") show"
        Print #outFN, posX + 600 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 50
        Print #outFN, posX + 1770 & " pontos -" & posY & " pontos moveto (" & Reg.TIPO_CLIENTE & ") show"
        posY = posY + 65
        Print #outFN, posX - 10 & " pontos -" & posY & " pontos moveto (" & Reg.CNPJ_CPF & ") show"
        Print #outFN, posX + 720 & " pontos -" & posY & " pontos moveto (" & Reg.RG & ") show"
        If Len(Trim(Reg.DATA_EMISSAO_RG)) > 0 Then
           Print #outFN, posX + 1730 & " pontos -" & posY & " pontos moveto (" & Mid(Reg.DATA_EMISSAO_RG, 1, 2) & "/" & Mid(Reg.DATA_EMISSAO_RG, 3, 2) & "/" & Mid(Reg.DATA_EMISSAO_RG, 5, 4) & ") show"
        End If
        posY = posY + 110
        Print #outFN, posX - 180 & " pontos -" & posY + 25 & " pontos moveto (" & Reg.ORG_EXPEDITOR & ") show"
        Print #outFN, posX + 120 & " pontos -" & posY & " pontos moveto (" & Reg.UF_ORG_EXPEDIDOR & ") show"
        If Len(Trim(Reg.DATA_NASCIMENTO)) > 0 Then
           Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Mid(Reg.DATA_NASCIMENTO, 1, 2) & "/" & Mid(Reg.DATA_NASCIMENTO, 3, 2) & "/" & Mid(Reg.DATA_NASCIMENTO, 5, 4) & ") show"
        End If
        Print #outFN, posX + 1150 & " pontos -" & posY & " pontos moveto (" & Reg.PROFISSAO & ") show"
        posY = posY + 50
        Print #outFN, posX + 1780 & " pontos -" & posY & " pontos moveto (" & Reg.SEXO & ") show"
        posY = posY + 65
        Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg.NOME_MAE & ") show"
        posY = posY + 50
        Print #outFN, posX + 1780 & " pontos -" & posY & " pontos moveto (" & Reg.ATIVO_INATIVO & ") show"
        posY = posY + 70
        Print #outFN, posX - 180 & " pontos -" & posY + 65 & " pontos moveto (" & Reg.TIPO_ENDERECO & ") show"
        Print #outFN, posX + 160 & " pontos -" & posY & " pontos moveto (" & Reg.COD_LOGRADOURO & ") show"
        Print #outFN, posX + 600 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        posY = posY + 115
        Print #outFN, posX - 120 & " pontos -" & posY & " pontos moveto (" & Reg.CEP & ") show"
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & ") show"
        Print #outFN, posX + 500 & " pontos -" & posY + 65 & " pontos moveto (" & Reg.REF & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.NUMERO & ") show"
        Print #outFN, posX + 1070 & " pontos -" & posY & " pontos moveto (" & Reg.COMPLEMENTO & ") show"
        posY = posY + 115
        Print #outFN, posX - 180 & " pontos -" & posY + 65 & " pontos moveto (" & Reg.TIPO_TELEFONE1 & ") show"
        Print #outFN, posX - 35 & " pontos -" & posY & " pontos moveto (" & Reg.DDD1 & ") show"
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg.TELEFONE1 & ") show"
        posY = posY + 315
        If Len(Trim(Reg.INSCRICAO)) > 0 Then
           Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Mid(Reg.INSCRICAO, 1, 3) & "." & Mid(Reg.INSCRICAO, 4, 3) & "." & Mid(Reg.INSCRICAO, 7, 3) & "." & Mid(Reg.INSCRICAO, 10, 4) & "." & Mid(Reg.INSCRICAO, 14, 3) & ") show"
        End If
        Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 1750 & " pontos -" & posY & " pontos moveto (" & Reg.COD_CLIENTE_PROPRIETARIO & ") show"
        posY = posY + 118
        Print #outFN, posX + 800 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.N_MORADORES & ") show"
        Print #outFN, posX + 1050 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.NOME_CONTA & ") show"
        Print #outFN, posX + 1680 & " pontos -" & posY & " pontos moveto (" & Reg.COD_CLIENTE_USUARIO & ") show"
        posY = posY + 120
        Print #outFN, posX + 160 & " pontos -" & posY & " pontos moveto (" & Reg.COD_LOGRADOURO_IMOVEL & ") show"
        Print #outFN, posX + 600 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO_IMOVEL & ") show"
        posY = posY + 115
        Print #outFN, posX - 120 & " pontos -" & posY & " pontos moveto (" & Reg.CEP_IMOVEL & ") show"
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO_IMOVEL & ") show"
        Print #outFN, posX + 500 & " pontos -" & posY + 65 & " pontos moveto (" & Reg.REF_IMOVEL & ") show"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.NUMERO_IMOVEL & ") show"
        Print #outFN, posX + 1070 & " pontos -" & posY & " pontos moveto (" & Reg.COMPLEMENTO_IMOVEL & ") show"
        posY = posY + 115
        Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg.CATEGORIA & ") show"
        posY = posY + 120
        Print #outFN, posX - 30 & " pontos -" & posY & " pontos moveto (" & Reg.QTD_ECONOMIA & ") show"
        posY = posY + 170
        Print #outFN, posX - 185 & " pontos -" & posY & " pontos moveto (" & Reg.AGUA & ") show"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.DIAM & ") show"
        Print #outFN, posX + 145 & " pontos -" & posY & " pontos moveto (" & Reg.MAT & ") show"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.INF & ") show"
        Print #outFN, posX + 490 & " pontos -" & posY & " pontos moveto (" & Reg.SUP & ") show"
        Print #outFN, posX + 680 & " pontos -" & posY & " pontos moveto (" & Reg.PISCINA & ") show"
        Print #outFN, posX + 840 & " pontos -" & posY & " pontos moveto (" & Reg.JARDIM & ") show"
        Print #outFN, posX + 1080 & " pontos -" & posY & " pontos moveto (" & Reg.PAVIMENTECAO & ") show"
        Print #outFN, posX + 1270 & " pontos -" & posY & " pontos moveto (" & Reg.RUA & ") show"
        Print #outFN, posX + 1450 & " pontos -" & posY & " pontos moveto (" & Reg.FONTE_ABASTECIMENTO & ") show"
        Print #outFN, posX + 1620 & " pontos -" & posY & " pontos moveto (" & Reg.POCO & ") show"
        Print #outFN, posX + 1760 & " pontos -" & posY - 35 & " pontos moveto (" & Reg.N_PONTOS & ") show"
        posY = posY + 120
        Print #outFN, posX - 185 & " pontos -" & posY & " pontos moveto (" & Reg.SIT_ESGOTO & ") show"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.DIA_ESGOTO & ") show"
        Print #outFN, posX + 145 & " pontos -" & posY & " pontos moveto (" & Reg.MAT_ESGOTO & ") show"
        Print #outFN, posX + 285 & " pontos -" & posY & " pontos moveto (" & Reg.PERFIL_IMOVEL & ") show"
        Print #outFN, posX + 585 & " pontos -" & posY & " pontos moveto (" & Reg.DESPEJO & ") show"
        Print #outFN, posX + 955 & " pontos -" & posY - 40 & " pontos moveto (" & Reg.LEITURA_INICIAL & ") show"
        Print #outFN, posX + 1135 & " pontos -" & posY - 5 & " pontos moveto (" & Reg.CAP & ") show"
        Print #outFN, posX + 1300 & " pontos -" & posY - 5 & " pontos moveto (" & Reg.MARCA & ") show"
        Print #outFN, posX + 1470 & " pontos -" & posY - 5 & " pontos moveto (" & Reg.LOCAL & ") show"
        Print #outFN, posX + 1655 & " pontos -" & posY - 5 & " pontos moveto (" & Reg.PROTECAO & ") show"
        Print #outFN, posX + 1790 & " pontos -" & posY - 5 & " pontos moveto (" & Reg.CAVALETE & ") show"
        
        posY = posY + 400
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

        'If wRec = 10 Then End
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
           
           Reg.COD_CLIENTE = Trim(Mid(linha, 288, 9))
           Reg.NOME = Trim(Mid(linha, 297, 23))
           Reg.TIPO_CLIENTE = Trim(Mid(linha, 320, 2))
           Reg.CNPJ_CPF = Trim(Mid(linha, 322, 14))
           Reg.RG = Trim(Mid(linha, 336, 13))
           Reg.DATA_EMISSAO_RG = Trim(Mid(linha, 349, 8))
           Reg.ORG_EXPEDITOR = Trim(Mid(linha, 357, 4))
           Reg.UF_ORG_EXPEDIDOR = Trim(Mid(linha, 361, 2))
           Reg.DATA_NASCIMENTO = Trim(Mid(linha, 363, 8))
           Reg.PROFISSAO = Trim(Mid(linha, 371, 18))
           Reg.SEXO = Trim(Mid(linha, 389, 1))
           Reg.NOME_MAE = Trim(Mid(linha, 390, 32))
           Reg.ATIVO_INATIVO = Trim(Mid(linha, 422, 1))
           Reg.TIPO_ENDERECO = Trim(Mid(linha, 423, 1))
           Reg.COD_LOGRADOURO = Trim(Mid(linha, 424, 9))
           Reg.ENDERECO = Trim(Mid(linha, 433, 60))
           Reg.CEP = Trim(Mid(linha, 493, 8))
           Reg.BAIRRO = Trim(Mid(linha, 501, 3))
           Reg.REF = Trim(Mid(linha, 504, 1))
           Reg.NUMERO = Trim(Mid(linha, 505, 5))
           Reg.COMPLEMENTO = Trim(Mid(linha, 510, 19))
           Reg.TIPO_TELEFONE1 = Trim(Mid(linha, 529, 1))
           Reg.DDD1 = Trim(Mid(linha, 530, 2))
           Reg.TELEFONE1 = Trim(Mid(linha, 532, 8))
           'Reg.RAMAL1 = Trim(Mid(linha, 540, 8))
           'Dados Imóvel
           Reg.INSCRICAO = Trim(Mid(linha, 568, 16))
           Reg.MATRICULA = Trim(Mid(linha, 584, 9))
           Reg.COD_CLIENTE_PROPRIETARIO = Trim(Mid(linha, 593, 12))
           Reg.N_MORADORES = Trim(Mid(linha, 621, 4))
           Reg.NOME_CONTA = Trim(Mid(linha, 625, 1))
           Reg.COD_CLIENTE_USUARIO = Trim(Mid(linha, 626, 12))
           Reg.COD_LOGRADOURO_IMOVEL = Trim(Mid(linha, 638, 9))
           Reg.ENDERECO_IMOVEL = Trim(Mid(linha, 647, 60))
           Reg.CEP_IMOVEL = Trim(Mid(linha, 707, 8))
           Reg.BAIRRO_IMOVEL = Trim(Mid(linha, 715, 4))
           Reg.REF_IMOVEL = Trim(Mid(linha, 268, 1))
           Reg.NUMERO_IMOVEL = Trim(Mid(linha, 269, 4))
           Reg.COMPLEMENTO_IMOVEL = Trim(Mid(linha, 724, 19))
           Reg.CATEGORIA = Trim(Mid(linha, 743, 2))
           Reg.QTD_ECONOMIA = Trim(Mid(linha, 745, 4))
           Reg.AGUA = Trim(Mid(linha, 791, 1))
           Reg.DIAM = Trim(Mid(linha, 792, 1))
           Reg.MAT = Trim(Mid(linha, 793, 1))
           Reg.INF = Trim(Mid(linha, 794, 1))
           Reg.SUP = Trim(Mid(linha, 795, 1))
           Reg.PISCINA = Trim(Mid(linha, 796, 1))
           Reg.JARDIM = Trim(Mid(linha, 797, 1))
           Reg.PAVIMENTECAO = Trim(Mid(linha, 798, 2))
           Reg.RUA = Trim(Mid(linha, 800, 2))
           Reg.FONTE_ABASTECIMENTO = Trim(Mid(linha, 802, 1))
           Reg.POCO = Trim(Mid(linha, 803, 1))
           Reg.N_PONTOS = Trim(Mid(linha, 805, 4))
           Reg.SIT_ESGOTO = Trim(Mid(linha, 809, 1))
           Reg.DIA_ESGOTO = Trim(Mid(linha, 810, 1))
           Reg.MAT_ESGOTO = Trim(Mid(linha, 810, 1))
           Reg.PERFIL_IMOVEL = Trim(Mid(linha, 811, 1))
           Reg.DESPEJO = Trim(Mid(linha, 812, 1))
           Reg.LEITURA_INICIAL = Trim(Mid(linha, 813, 6))
           Reg.CAP = Trim(Mid(linha, 819, 2))
           Reg.MARCA = Trim(Mid(linha, 821, 2))
           Reg.LOCAL = Trim(Mid(linha, 823, 2))
           Reg.PROTECAO = Trim(Mid(linha, 825, 1))
           Reg.CAVALETE = Trim(Mid(linha, 826, 1))

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
        Open App.Path & "\Boletim_Cadastral_Compesa_A4.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Boletim Cadastral Compesa A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        'Dados Cliente
        Reg.COD_CLIENTE = ""
        Reg.NOME = ""
        Reg.TIPO_CLIENTE = ""
        Reg.CNPJ_CPF = ""
        Reg.RG = ""
        Reg.DATA_EMISSAO_RG = ""
        Reg.ORG_EXPEDITOR = ""
        Reg.UF_ORG_EXPEDIDOR = ""
        Reg.PROFISSAO = ""
        Reg.SEXO = ""
        Reg.NOME_MAE = ""
        Reg.ATIVO_INATIVO = ""
        Reg.DATA_NASCIMENTO = ""
        Reg.TIPO_ENDERECO = ""
        Reg.COD_LOGRADOURO = ""
        Reg.ENDERECO = ""
        Reg.COMPLEMENTO = ""
        Reg.CEP = ""
        Reg.BAIRRO = ""
        Reg.REF = ""
        Reg.NUMERO = ""
        Reg.TIPO_TELEFONE1 = ""
        Reg.DDD1 = ""
        Reg.TELEFONE1 = ""
        Reg.RAMAL1 = ""
        Reg.TIPO_TELEFONE2 = ""
        Reg.DDD2 = ""
        Reg.TELEFONE2 = ""
        Reg.RAMAL2 = ""
        'Dados Imóvel
        Reg.INSCRICAO = ""
        Reg.MATRICULA = ""
        Reg.N_MORADORES = ""
        Reg.NOME_CONTA = ""
        Reg.COD_CLIENTE_PROPRIETARIO = ""
        Reg.COD_CLIENTE_USUARIO = ""
        Reg.COD_LOGRADOURO_IMOVEL = ""
        Reg.ENDERECO_IMOVEL = ""
        Reg.COMPLEMENTO_IMOVEL = ""
        Reg.CEP_IMOVEL = ""
        Reg.BAIRRO_IMOVEL = ""
        Reg.REF_IMOVEL = ""
        Reg.NUMERO_IMOVEL = ""
        Reg.CATEGORIA = ""
        Reg.QTD_ECONOMIA = ""
        Reg.AGUA = ""
        Reg.DIAM = ""
        Reg.MAT = ""
        Reg.JARDIM = ""
        Reg.PAVIMENTECAO = ""
        Reg.RUA = ""
        Reg.FONTE_ABASTECIMENTO = ""
        Reg.N_PONTOS = ""
        Reg.SIT_ESGOTO = ""
        Reg.PERFIL_IMOVEL = ""
        Reg.DESPEJO = ""
End Sub
