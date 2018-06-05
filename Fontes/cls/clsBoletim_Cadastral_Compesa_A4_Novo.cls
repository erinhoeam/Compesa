VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoletim_Cad_Compesa_A4_Novo"
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

Const AppName = "COMPESA", AppDesc = "Boletim Cadastral Compesa A4 Novo"

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
        ANT_NOME As String
        NOME As String
        CNPJ_CPF As String
        RG As String
        DATA_EMISSAO_RG As String
        UF_ORG_EXPEDIDOR As String
        SEXO As String
        NOME_MAE As String
        TIPO_ENDERECO As String
        ENDERECO As String
        COMPLEMENTO As String
        CEP As String
        CIDADE As String
        BAIRRO As String
        NUMERO As String
        TIPO_TELEFONE1 As String
        DDD1 As String
        TELEFONE1 As String
        RAMAL1 As String
        'Dados Usuário
        ANT_NOMEUSUARIO As String
        NOMEUSUARIO As String
        CPF_CNPJ_USER As String
        RG_USER As String
        UF_USER As String
        SEXO_USER As String
        TIPO_FONE_USER As String
        DDD_FONE_FONE_USER As String
        FONE1_USER As String
        FONE2_USER As String
        RAMAL_USER As String
        'Dados Imóvel
        COD_LOGRADOURO As String
        INSCRICAO As String
        MATRICULA As String
        ENDERECO_IMOVEL As String
        COMPLEMENTO_IMOVEL As String
        CEP_IMOVEL As String
        BAIRRO_IMOVEL As String
        CIDADE_IMOVEL As String
        REF_IMOVEL As String
        NUMERO_IMOVEL As String
        N_MORADORES As String
        N_MEDIDOR_CELPE As String
        N_HIDROMETRO As String
        SUBCATEGORIA1 As String
        SUBCATEGORIA2 As String
        SUBCATEGORIA3 As String
        SUBCATEGORIA4 As String
        SUBCATEGORIA5 As String
        SUBCATEGORIA6 As String
        ECONOMIA1 As String
        ECONOMIA2 As String
        ECONOMIA3 As String
        ECONOMIA4 As String
        ECONOMIA5 As String
        ECONOMIA6 As String
        SIT_AGUA As String
        PAVIMENTO_CALCADA As String
        PAVIMENTO_RUA As String
        FONTE_ABASTECIMENTO As String
        SITUACAO_ESGOTO As String
        PERFIL_IMOVEL As String
        CAPACIDADE_HIDROMETRO As String
        MARCA_HIDROMETRO As String
        LOCAL_ISNT_HIDROMETRO As String
        TIPO_PROTECAO As String
        CAVALETE As String
        RENDA_COMPROVADA As String
        RENDA_DECLARADA As String
        N_BOLSAFAMILIA As String
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
        posY = 550
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        'Rota
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 430 & " pontos -" & posY & " pontos moveto (" & Reg.ANT_NOME & " " & Reg.NOME & ") show"
        Print #outFN, posX + 1730 & " pontos -" & posY - 320 & " pontos moveto (" & Format(Date, "dd/MM/yyyy") & ") show"
        Print #outFN, posX + 1730 & " pontos -" & posY - 215 & " pontos moveto (" & Format(wRec, "000000") & ") show"
        If Reg.SEXO = "1" Then
           Print #outFN, posX + 1850 & " pontos -" & posY & " pontos moveto (" & Reg.SEXO & ") show"
        ElseIf Reg.SEXO = "2" Then
           Print #outFN, posX + 1965 & " pontos -" & posY & " pontos moveto (" & Reg.SEXO & ") show"
        End If
        posY = posY + 90
        Print #outFN, posX + 100 & " pontos -" & posY & " pontos moveto (" & Reg.CNPJ_CPF & ") show"
        Print #outFN, posX + 755 & " pontos -" & posY & " pontos moveto (" & Reg.RG & ") show"
        Print #outFN, posX + 1215 & " pontos -" & posY & " pontos moveto (" & Reg.UF_ORG_EXPEDIDOR & ") show"
        Print #outFN, posX + 1515 & " pontos -" & posY & " pontos moveto (" & Reg.DDD1 & " " & Reg.TELEFONE1 & ") show"
        Print #outFN, posX + 1970 & " pontos -" & posY & " pontos moveto (" & Reg.TIPO_TELEFONE1 & ") show"
        posY = posY + 90
        Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        'Print #outFN, posX + 1905 & " pontos -" & posY & " pontos moveto (" & Reg.NUMERO & ") show"
        posY = posY + 90
        Print #outFN, posX + 190 & " pontos -" & posY & " pontos moveto (" & Reg.COMPLEMENTO & ") show"
        Print #outFN, posX + 1180 & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & ") show"
        posY = posY + 90
        Print #outFN, posX + 120 & " pontos -" & posY & " pontos moveto (" & Reg.CIDADE & ") show"
        Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg.CEP & ") show"
        Print #outFN, posX + 1965 & " pontos -" & posY & " pontos moveto (" & Reg.TIPO_ENDERECO & ") show"
        posY = posY + 90
        Print #outFN, posX + 360 & " pontos -" & posY & " pontos moveto (" & Reg.ANT_NOMEUSUARIO & " " & Reg.NOMEUSUARIO & ") show"
        If Reg.SEXO_USER = "1" Then
           Print #outFN, posX + 1850 & " pontos -" & posY & " pontos moveto (" & Reg.SEXO_USER & ") show"
        ElseIf Reg.SEXO_USER = "2" Then
           Print #outFN, posX + 1965 & " pontos -" & posY & " pontos moveto (" & Reg.SEXO_USER & ") show"
        End If
        posY = posY + 90
        Print #outFN, posX + 100 & " pontos -" & posY & " pontos moveto (" & Reg.CPF_CNPJ_USER & ") show"
        Print #outFN, posX + 755 & " pontos -" & posY & " pontos moveto (" & Reg.RG_USER & ") show"
        Print #outFN, posX + 1215 & " pontos -" & posY & " pontos moveto (" & Reg.UF_USER & ") show"
        Print #outFN, posX + 1515 & " pontos -" & posY & " pontos moveto (" & Reg.DDD_FONE_FONE_USER & " " & Reg.FONE1_USER & ") show"
        Print #outFN, posX + 1970 & " pontos -" & posY & " pontos moveto (" & Reg.TIPO_FONE_USER & ") show"
        posY = posY + 290
        Print #outFN, posX + 100 & " pontos -" & posY & " pontos moveto (" & Mid(Reg.INSCRICAO, 1, 3) & "." & Mid(Reg.INSCRICAO, 4, 3) & "." & Mid(Reg.INSCRICAO, 7, 3) & "." & Mid(Reg.INSCRICAO, 10, 4) & "." & Mid(Reg.INSCRICAO, 14, 3) & ") show"
        Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (" & Mid(Reg.MATRICULA, 1, 8) & "-" & Mid(Reg.MATRICULA, 9, 1) & ") show"
        Print #outFN, posX + 1965 & " pontos -" & posY & " pontos moveto (" & Reg.PERFIL_IMOVEL & ") show"
        posY = posY + 90
        Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO_IMOVEL & " " & Reg.NUMERO_IMOVEL & ") show"
        'Print #outFN, posX + 1905 & " pontos -" & posY & " pontos moveto (" & Reg.NUMERO_IMOVEL & ") show"
        posY = posY + 90
        Print #outFN, posX + 190 & " pontos -" & posY & " pontos moveto (" & Reg.COMPLEMENTO_IMOVEL & ") show"
        Print #outFN, posX + 1180 & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO_IMOVEL & ") show"
        posY = posY + 90
        Print #outFN, posX + 120 & " pontos -" & posY & " pontos moveto (" & Reg.CIDADE_IMOVEL & ") show"
        Print #outFN, posX + 870 & " pontos -" & posY & " pontos moveto (" & Reg.CEP_IMOVEL & ") show"
        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.COD_LOGRADOURO & ") show"
        
        posY = posY + 190
        Print #outFN, posX + 380 & " pontos -" & posY & " pontos moveto (" & Reg.SUBCATEGORIA1 & ") show"
        Print #outFN, posX + 560 & " pontos -" & posY & " pontos moveto (" & Reg.SUBCATEGORIA2 & ") show"
        Print #outFN, posX + 750 & " pontos -" & posY & " pontos moveto (" & Reg.SUBCATEGORIA3 & ") show"
        Print #outFN, posX + 940 & " pontos -" & posY & " pontos moveto (" & Reg.SUBCATEGORIA4 & ") show"
        Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg.SUBCATEGORIA5 & ") show"
        Print #outFN, posX + 1320 & " pontos -" & posY & " pontos moveto (" & Reg.SUBCATEGORIA6 & ") show"
        posY = posY + 90
        Print #outFN, posX + 360 & " pontos -" & posY & " pontos moveto (" & Reg.ECONOMIA1 & ") show"
        Print #outFN, posX + 540 & " pontos -" & posY & " pontos moveto (" & Reg.ECONOMIA2 & ") show"
        Print #outFN, posX + 730 & " pontos -" & posY & " pontos moveto (" & Reg.ECONOMIA3 & ") show"
        Print #outFN, posX + 920 & " pontos -" & posY & " pontos moveto (" & Reg.ECONOMIA4 & ") show"
        Print #outFN, posX + 1110 & " pontos -" & posY & " pontos moveto (" & Reg.ECONOMIA5 & ") show"
        Print #outFN, posX + 1300 & " pontos -" & posY & " pontos moveto (" & Reg.ECONOMIA6 & ") show"
        posY = posY + 90
        Print #outFN, posX + 250 & " pontos -" & posY & " pontos moveto (" & Reg.N_MORADORES & ") show"
        Print #outFN, posX + 750 & " pontos -" & posY & " pontos moveto (" & Reg.N_MEDIDOR_CELPE & ") show"
        If Reg.PAVIMENTO_CALCADA = "00" Then
           Print #outFN, posX + 1540 & " pontos -" & posY & " pontos moveto (" & Reg.PAVIMENTO_CALCADA & ") show"
        ElseIf Reg.PAVIMENTO_CALCADA = "01" Then
           Print #outFN, posX + 1780 & " pontos -" & posY & " pontos moveto (" & Reg.PAVIMENTO_CALCADA & ") show"
        End If
        posY = posY + 90
        If Reg.FONTE_ABASTECIMENTO = "1" Then
           Print #outFN, posX + 1180 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.FONTE_ABASTECIMENTO = "2" Then
           Print #outFN, posX + 900 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.FONTE_ABASTECIMENTO = "3" Then
           Print #outFN, posX + 1540 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.FONTE_ABASTECIMENTO = "4" Then
           Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        End If
        posY = posY + 280
        If Reg.SITUACAO_ESGOTO = "1" Then
           Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SITUACAO_ESGOTO = "2" Then
           Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SITUACAO_ESGOTO = "3" Then
           Print #outFN, posX + 850 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SITUACAO_ESGOTO = "6" Then
           Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SITUACAO_ESGOTO = "5" Then
           Print #outFN, posX + 1490 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        End If
        posY = posY + 90
        If Reg.SIT_AGUA = "1" Then
           Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SIT_AGUA = "2" Then
           Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SIT_AGUA = "3" Then
           Print #outFN, posX + 850 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SIT_AGUA = "4" Then
           Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SIT_AGUA = "5" Then
           Print #outFN, posX + 1490 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        ElseIf Reg.SIT_AGUA = "6" Then
           Print #outFN, posX + 1770 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        End If
        posY = posY + 90
        If Len(Trim(Reg.N_HIDROMETRO)) > 0 Then
           Print #outFN, posX + 560 & " pontos -" & posY & " pontos moveto (" & "X" & ") show"
        End If
        Print #outFN, posX + 890 & " pontos -" & posY & " pontos moveto (" & Reg.N_HIDROMETRO & ") show"
        
        Print #outFN, posX + 1580 & " pontos -" & posY & " pontos moveto (" & Reg.MARCA_HIDROMETRO & ") show"
        Print #outFN, posX + 1960 & " pontos -" & posY & " pontos moveto (" & Reg.CAPACIDADE_HIDROMETRO & ") show"
        posY = posY + 90
        Print #outFN, posX + 370 & " pontos -" & posY & " pontos moveto (" & Reg.LOCAL_ISNT_HIDROMETRO & ") show"
        Print #outFN, posX + 920 & " pontos -" & posY & " pontos moveto (" & Reg.TIPO_PROTECAO & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Reg.CAVALETE & ") show"
        posY = posY + 170
        If Trim(Reg.RENDA_COMPROVADA) = "COMPROVADA" Then
           Print #outFN, posX + 450 & " pontos -" & posY & " pontos moveto (" & Space(10 - Len(Trim(Reg.RENDA_DECLARADA))) & Trim(Reg.RENDA_DECLARADA) & ") show"
        ElseIf Trim(Reg.RENDA_COMPROVADA) = "DECLARADA" Then
           Print #outFN, posX + 1025 & " pontos -" & posY & " pontos moveto (" & Space(10 - Len(Trim(Reg.RENDA_DECLARADA))) & Trim(Reg.RENDA_DECLARADA) & ") show"
        ElseIf Trim(Reg.RENDA_COMPROVADA) = "NAO INFORMADA" Then
           Print #outFN, posX + 600 & " pontos -" & posY & " pontos moveto (" & Space(10 - Len(Trim(Reg.RENDA_DECLARADA))) & Trim(Reg.RENDA_DECLARADA) & ") show"
        End If
        Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg.N_BOLSAFAMILIA & ") show"
        posY = posY + 780
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

        'If wRec = 200 Then End
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

           Reg.ANT_NOME = Trim(Mid(linha, 9, 9))
           Reg.NOME = Trim(Mid(linha, 18, 23))
           Reg.CNPJ_CPF = Trim(Mid(linha, 43, 14))
           Reg.RG = Trim(Mid(linha, 57, 13))
           Reg.UF_ORG_EXPEDIDOR = Trim(Mid(linha, 82, 2))
           Reg.SEXO = Trim(Mid(linha, 110, 1))
           Reg.TIPO_ENDERECO = Trim(Mid(linha, 145, 1))
           Reg.ENDERECO = Trim(Mid(linha, 154, 60))
           Reg.CEP = Trim(Mid(linha, 214, 8))
           Reg.BAIRRO = Trim(Mid(linha, 892, 30))
           Reg.CIDADE = Trim(Mid(linha, 922, 30))
           Reg.NUMERO = Trim(Mid(linha, 226, 5))
           Reg.COMPLEMENTO = Trim(Mid(linha, 231, 19))
           Reg.TIPO_TELEFONE1 = Trim(Mid(linha, 250, 1))
           Reg.DDD1 = Trim(Mid(linha, 251, 2))
           Reg.TELEFONE1 = Trim(Mid(linha, 253, 8))
           Reg.RAMAL1 = Trim(Mid(linha, 261, 4))
           'Dados Usuario
           Reg.ANT_NOMEUSUARIO = Trim(Mid(linha, 288, 9))
           Reg.NOMEUSUARIO = Trim(Mid(linha, 297, 23))
           Reg.CPF_CNPJ_USER = Trim(Mid(linha, 322, 14))
           Reg.RG_USER = Trim(Mid(linha, 336, 13))
           Reg.UF_USER = Trim(Mid(linha, 361, 2))
           Reg.SEXO_USER = Trim(Mid(linha, 389, 1))
           Reg.TIPO_FONE_USER = Trim(Mid(linha, 529, 1))
           Reg.DDD_FONE_FONE_USER = Trim(Mid(linha, 530, 2))
           Reg.FONE1_USER = Trim(Mid(linha, 532, 8))
           Reg.RAMAL_USER = Trim(Mid(linha, 540, 4))
           'Dados Imóvel
           Reg.INSCRICAO = Trim(Mid(linha, 568, 16))
           Reg.MATRICULA = Trim(Mid(linha, 584, 9))
           Reg.ENDERECO_IMOVEL = Trim(Mid(linha, 647, 60))
           Reg.CEP_IMOVEL = Trim(Mid(linha, 707, 8))
           Reg.BAIRRO_IMOVEL = Trim(Mid(linha, 952, 30))
           Reg.CIDADE_IMOVEL = Trim(Mid(linha, 982, 30))
           Reg.NUMERO_IMOVEL = Trim(Mid(linha, 505, 5))
           Reg.COMPLEMENTO_IMOVEL = Trim(Mid(linha, 724, 19))
           Reg.N_MORADORES = Trim(Mid(linha, 621, 4))
           Reg.COD_LOGRADOURO = Trim(Mid(linha, 638, 9))
           Reg.N_MEDIDOR_CELPE = Trim(Mid(linha, 1022, 10))
           Reg.N_HIDROMETRO = Trim(Mid(linha, 1012, 10))
           Reg.SUBCATEGORIA1 = Trim(Mid(linha, 743, 2))
           Reg.SUBCATEGORIA2 = Trim(Mid(linha, 749, 2))
           Reg.SUBCATEGORIA3 = Trim(Mid(linha, 755, 2))
           Reg.SUBCATEGORIA4 = Trim(Mid(linha, 761, 2))
           Reg.SUBCATEGORIA5 = Trim(Mid(linha, 767, 2))
           Reg.SUBCATEGORIA6 = Trim(Mid(linha, 773, 2))
           Reg.ECONOMIA1 = Trim(Mid(linha, 745, 4))
           Reg.ECONOMIA2 = Trim(Mid(linha, 751, 4))
           Reg.ECONOMIA3 = Trim(Mid(linha, 757, 4))
           Reg.ECONOMIA4 = Trim(Mid(linha, 763, 4))
           Reg.ECONOMIA5 = Trim(Mid(linha, 769, 4))
           Reg.ECONOMIA6 = Trim(Mid(linha, 775, 4))
           Reg.SIT_AGUA = Trim(Mid(linha, 791, 1))
           Reg.PAVIMENTO_CALCADA = Trim(Mid(linha, 798, 2))
           Reg.PAVIMENTO_RUA = Trim(Mid(linha, 800, 2))
           Reg.FONTE_ABASTECIMENTO = Trim(Mid(linha, 802, 1))
           Reg.SITUACAO_ESGOTO = Trim(Mid(linha, 808, 1))
           Reg.PERFIL_IMOVEL = Trim(Mid(linha, 811, 1))
           Reg.CAPACIDADE_HIDROMETRO = Trim(Mid(linha, 819, 2))
           Reg.MARCA_HIDROMETRO = Trim(Mid(linha, 821, 2))
           Reg.LOCAL_ISNT_HIDROMETRO = Trim(Mid(linha, 823, 2))
           Reg.TIPO_PROTECAO = Trim(Mid(linha, 825, 1))
           Reg.CAVALETE = Trim(Mid(linha, 826, 1))
           Reg.RENDA_COMPROVADA = Trim(Mid(linha, 1040, 20))
           Reg.RENDA_DECLARADA = Trim(Mid(linha, 1060, 10))
           Reg.N_BOLSAFAMILIA = Trim(Mid(linha, 1070, 31))

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
        Open App.Path & "\Boletim_Cadastral_Novo.ps" For Input As #200
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
        With Reg
            'Dados Cliente
            .NOME = ""
            .CNPJ_CPF = ""
            .RG = ""
            .DATA_EMISSAO_RG = ""
            .UF_ORG_EXPEDIDOR = ""
            .SEXO = ""
            .NOME_MAE = ""
            .TIPO_ENDERECO = ""
            .ENDERECO = ""
            .COMPLEMENTO = ""
            .CEP = ""
            .BAIRRO = ""
            .CIDADE = ""
            .NUMERO = ""
            .TIPO_TELEFONE1 = ""
            .DDD1 = ""
            .TELEFONE1 = ""
            .RAMAL1 = ""
            'Dados Usuário
            .NOMEUSUARIO = ""
            .CPF_CNPJ_USER = ""
            .RG_USER = ""
            .UF_USER = ""
            .SEXO_USER = ""
            .TIPO_FONE_USER = ""
            .DDD_FONE_FONE_USER = ""
            .FONE1_USER = ""
            .FONE2_USER = ""
            .RAMAL_USER = ""
            'Dados Imóvel
            .INSCRICAO = ""
            .MATRICULA = ""
            .ENDERECO_IMOVEL = ""
            .COMPLEMENTO_IMOVEL = ""
            .CEP_IMOVEL = ""
            .BAIRRO_IMOVEL = ""
            .CIDADE_IMOVEL = ""
            .REF_IMOVEL = ""
            .NUMERO_IMOVEL = ""
            .SUBCATEGORIA1 = ""
            .SUBCATEGORIA2 = ""
            .SUBCATEGORIA3 = ""
            .SUBCATEGORIA4 = ""
            .SUBCATEGORIA5 = ""
            .SUBCATEGORIA6 = ""
            .ECONOMIA1 = ""
            .ECONOMIA2 = ""
            .ECONOMIA3 = ""
            .ECONOMIA4 = ""
            .ECONOMIA5 = ""
            .ECONOMIA6 = ""
            .SIT_AGUA = ""
            .PAVIMENTO_CALCADA = ""
            .PAVIMENTO_RUA = ""
            .FONTE_ABASTECIMENTO = ""
            .SITUACAO_ESGOTO = ""
            .PERFIL_IMOVEL = ""
            .CAPACIDADE_HIDROMETRO = ""
            .MARCA_HIDROMETRO = ""
            .LOCAL_ISNT_HIDROMETRO = ""
            .CAVALETE = ""
            .TIPO_PROTECAO = ""
            .RENDA_COMPROVADA = ""
            .RENDA_DECLARADA = ""
            .N_BOLSAFAMILIA = ""
        End With
End Sub
