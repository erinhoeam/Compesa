VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConta_Boleto_Compesa_A4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Ericsson Alves de Melo
'                 Analista de Sistemas
'Data: 27 de Novembro de 2008
'Contatos - Email: ericsson@sistemap.com.br
'                  erinhoeam@hotmail.com
'           Fone: (81) 9985.9131
'Aplicação: Conta Compesa Boleto A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Conta Compesa Boleto A4"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        TIPO_CONTA As String
        ESCRITORIO As String
        MATRICULA As String
        CLIENTE As String
        VENCIMENTO1 As String
        VENCIMENTO_FICHA As String
        INSCRICAO As String
        ENDERECO As String
        ENDERECO1 As String
        FATURA As String
        RESPONSAVEL As String
        ENDERECO_ENTREGA As String
        ENDERECO_ENTREGA1 As String
        CEP As String
        AGUA As String
        ESGOTO As String
        ULTIMOS_CONSUMOS(3) As String
        LEITURA_ANTERIOR As String
        LEITURA_ATUAL As String
        CONSUMO_M3 As String
        CONSUMO_DIAS As String
        CONSUMO_MEDIA As String
        TYPE_CONSUMO1 As String
        TYPE_CONSUMO2 As String
        COSUMO_ESGOTO As String
        PERIODO1 As String
        PERIODO2 As String
        ECONOMIA As String
        CONSUMO_POR_ECONOMIA As String
        COD_AUXILIAR As String
        DESCRICAO(15) As String
        FAIXA(15) As String
        VALOR(15) As String
        TOTAL As String
        MENSAGEM(3) As String
        GERENCIA_REGIONAL As String
        MES_ANO As String
        TURBIDEZ As String
        CLORO As String
        SEQ As String
        LINHA_DIGITAVEL As String
        LOCAL_PAGAMENTO As String
        DATA_DOCUMENTO As String
        N_DOCUMENTO As String
        ESPECIE_DOC As String
        ACEITE As String
        DATA_PROCESSAMENTO As String
        NOSSO_NUMERO As String
        CEDENTE As String
        VENCIMENTO_BOLETO As String
        AGENCIA_COD_CEDENTE As String
        CARTEIRA As String
        ESPECIE_MOEDA As String
        BAR_CODE As String
        NAO_RECEBER As String
        GRUPO As String
        FIRMA As String
        VALOR_TOTAL As String
        Linha1_SACADO As String
        Linha2_SACADO As String
End Type
'Variáveis gerais para os tipos
Private Reg As TipoReg

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
        posX = 215
        posY = 380
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        'Rota
        Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
        Print #outFN, posX + 300 & " pontos -" & posY - 280 & " pontos moveto (" & Reg.TIPO_CONTA & ") show"
        Print #outFN, posX + 350 & " pontos -" & posY - 50 & " pontos moveto (Escritório: " & Reg.ESCRITORIO & ") show"
        posY = posY + 50
        'Nome, Matricula
        Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
        Print #outFN, posX + 380 & " pontos -" & posY & " pontos moveto (" & Reg.CLIENTE & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 1770 & " pontos -" & posY - 30 & " pontos moveto (" & Reg.VENCIMENTO1 & ") show"
        posY = posY + 120
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX - 15 & " pontos -" & posY - 25 & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 465 & " pontos -" & posY - 30 & " pontos moveto (" & Reg.ENDERECO & ") show"
        Print #outFN, posX + 465 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO1 & ") show"
        Print #outFN, posX + 1775 & " pontos -" & posY - 25 & " pontos moveto (" & Reg.FATURA & ") show"
        posY = posY + 80
        Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
        Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg.RESPONSAVEL & ") show"
        Print #outFN, posX + 170 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO_ENTREGA & ") show"
        Print #outFN, posX + 170 & " pontos -" & posY + 35 & " pontos moveto (" & Reg.ENDERECO_ENTREGA1 & ") show"
        Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg.AGUA & ") show"
        Print #outFN, posX + 1705 & " pontos -" & posY & " pontos moveto (" & Reg.ESGOTO & ") show"
        posY = posY + 130
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 8 scalefont setfont"
        For i2 = 1 To 3
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.ULTIMOS_CONSUMOS(i2) & ") show"
            If i2 = 2 Then
               Print #outFN, posX + 1045 & " pontos -" & posY & " pontos moveto (" & Reg.LEITURA_ANTERIOR & ") show"
               Print #outFN, posX + 1360 & " pontos -" & posY & " pontos moveto (" & Reg.LEITURA_ATUAL & ") show"
               Print #outFN, posX + 1620 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_M3 & ") show"
               Print #outFN, posX + 1730 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_DIAS & ") show"
               Print #outFN, posX + 1880 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_MEDIA & ") show"
               Print #outFN, posX + 1045 & " pontos -" & posY + 40 & " pontos moveto (" & Reg.PERIODO1 & ") show"
               Print #outFN, posX + 1360 & " pontos -" & posY + 40 & " pontos moveto (" & Reg.PERIODO2 & ") show"
            End If
            posY = posY + 40
        Next
        posY = posY + 20
        Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg.ECONOMIA & ") show"
        Print #outFN, posX + 205 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_POR_ECONOMIA & ") show"
        Print #outFN, posX + 510 & " pontos -" & posY & " pontos moveto (" & Reg.COD_AUXILIAR & ") show"
        Print #outFN, posX + 1010 & " pontos -" & posY - 30 & " pontos moveto (" & Reg.TYPE_CONSUMO1 & ") show"
        Print #outFN, posX + 1550 & " pontos -" & posY - 30 & " pontos moveto (" & Reg.TYPE_CONSUMO2 & ") show"
        Print #outFN, posX + 1010 & " pontos -" & posY + 20 & " pontos moveto (" & Reg.COSUMO_ESGOTO & ") show"
        posY = posY + 150
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7.5 scalefont setfont"
        For i3 = 1 To 15
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.DESCRICAO(i3) & ") show"
            Print #outFN, posX + 1320 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FAIXA(i3))) & Reg.FAIXA(i3) & ") show"
            Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(30 - Len(Reg.VALOR(i3))) & Reg.VALOR(i3) & ") show"

            posY = posY + 35
        Next
        posY = 1650
       'Total Débito
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1470 & " pontos -" & posY & " pontos moveto (" & Space(30 - Len(Reg.TOTAL)) & Reg.TOTAL & ") show"
        posY = posY + 80
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
        For i4 = 1 To 3
            Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg.MENSAGEM(i4) & ") show"
            
            posY = posY + 30
        Next
        posY = posY + 130
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 350 & " pontos -" & posY & " pontos moveto (" & Reg.GERENCIA_REGIONAL & ") show"
        Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg.MES_ANO & ") show"
        posY = posY + 80
        Print #outFN, posX + 900 & " pontos -" & posY & " pontos moveto (" & Reg.TURBIDEZ & ") show"
        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.CLORO & ") show"
        posY = posY + 55
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO & ") show"
        Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg.FIRMA & ") show"
        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (" & Reg.SEQ & ") show"
        posY = posY + 260
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 12 scalefont setfont"
        Print #outFN, posX + 630 & " pontos -" & posY & " pontos moveto (" & Reg.LINHA_DIGITAVEL & ") show"
        posY = posY + 80
        Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
        Print #outFN, posX - 10 & " pontos -" & posY & " pontos moveto (" & Reg.LOCAL_PAGAMENTO & ") show"
        Print #outFN, posX + 1460 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO_FICHA & ") show"
        posY = posY + 70
        Print #outFN, posX - 10 & " pontos -" & posY & " pontos moveto (" & Reg.CEDENTE & ") show"
        Print #outFN, posX + 1460 & " pontos -" & posY + 15 & " pontos moveto (" & Reg.AGENCIA_COD_CEDENTE & ") show"
        posY = posY + 75
        Print #outFN, posX - 10 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_DOCUMENTO & ") show"
        Print #outFN, posX + 340 & " pontos -" & posY & " pontos moveto (" & Reg.N_DOCUMENTO & ") show"
        Print #outFN, posX + 720 & " pontos -" & posY & " pontos moveto (" & Reg.ESPECIE_DOC & ") show"
        Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg.ACEITE & ") show"
        Print #outFN, posX + 1095 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_PROCESSAMENTO & ") show"
        Print #outFN, posX + 1460 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.NOSSO_NUMERO & ") show"
        posY = posY + 75
        Print #outFN, posX + 340 & " pontos -" & posY & " pontos moveto (" & Reg.CARTEIRA & ") show"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & "R$" & ") show"
        posY = posY + 75
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1180 & " pontos -" & posY & " pontos moveto (" & Space(30 - Len(Reg.TOTAL)) & Reg.TOTAL & ") show"
        posY = posY + 345
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 60 & " pontos -" & posY & " pontos moveto (" & Reg.Linha1_SACADO & ") show"
        posY = posY + 35
        Print #outFN, posX + 60 & " pontos -" & posY & " pontos moveto (" & Reg.Linha2_SACADO & ") show"
        Print #outFN, posX + 60 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.ENDERECO1 & ") show"
        posY = posY + 320
        Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
        Print #outFN, posX + 110 & " pontos -" & posY & " pontos moveto (" & Reg.BAR_CODE & ") show"
        posY = posY + 35
        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 1750 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Print #outFN, "showpage"
        'Fim da Parte Interna
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

'       If wRec = 1 Then End
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

        frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        frmPrincipal.ProgressBarProcess.Value = 0.0001
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha
           
           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")
           
           Reg.TIPO_CONTA = Trim(Mid(linha, 1, 50))
           Reg.ESCRITORIO = Trim(Mid(linha, 51, 30))
           Reg.MATRICULA = Trim(Mid(linha, 81, 9))
           Reg.CLIENTE = Trim(Mid(linha, 90, 30))
           Reg.VENCIMENTO1 = Trim(Mid(linha, 120, 10))
           Reg.INSCRICAO = Trim(Mid(linha, 130, 20))
           
           Reg.ENDERECO = Trim(Mid(linha, 150, 50))
           Reg.ENDERECO1 = Trim(Mid(linha, 2647, 50))
           Reg.FATURA = Trim(Mid(linha, 200, 9))
           Reg.RESPONSAVEL = Trim(Mid(linha, 208, 9))
           Reg.ENDERECO_ENTREGA = Trim(Mid(linha, 217, 50))
           Reg.ENDERECO_ENTREGA1 = Trim(Mid(linha, 2693, 50))
           Reg.AGUA = Trim(Mid(linha, 267, 20))
           Reg.ESGOTO = Trim(Mid(linha, 287, 20))

           For i1 = 1 To 3
               If i1 = 1 Then
                  Reg.ULTIMOS_CONSUMOS(i1) = Trim(Mid(linha, 307, 38))
               End If
               If i1 = 2 Then
                  Reg.ULTIMOS_CONSUMOS(i1) = Trim(Mid(linha, 345, 38))
               End If
               If i1 = 3 Then
                  Reg.ULTIMOS_CONSUMOS(i1) = Trim(Mid(linha, 410, 38))
               End If
           Next
           
           Reg.LEITURA_ANTERIOR = Trim(Mid(linha, 383, 7))
           Reg.LEITURA_ATUAL = Trim(Mid(linha, 390, 7))
           Reg.CONSUMO_M3 = Trim(Mid(linha, 397, 5))
           Reg.CONSUMO_DIAS = Trim(Mid(linha, 402, 2))
           Reg.CONSUMO_MEDIA = Trim(Mid(linha, 404, 6))
           Reg.PERIODO1 = Trim(Mid(linha, 448, 10))
           Reg.PERIODO2 = Trim(Mid(linha, 458, 10))
           Reg.TYPE_CONSUMO1 = Trim(Mid(linha, 468, 28))
           Reg.TYPE_CONSUMO2 = Trim(Mid(linha, 496, 26))
           Reg.ECONOMIA = Trim(Mid(linha, 522, 2))
           Reg.CONSUMO_POR_ECONOMIA = Trim(Mid(linha, 524, 5))
           Reg.COD_AUXILIAR = Trim(Mid(linha, 529, 16))
           Reg.COSUMO_ESGOTO = Trim(Mid(linha, 545, 31))
           
           For i2 = 1 To 15
               If i2 = 1 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 576, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 626, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 641, 14))
               End If
               If i2 = 2 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 655, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 705, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 720, 14))
               End If
               If i2 = 3 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 734, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 784, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 799, 14))
               End If
               If i2 = 4 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 813, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 862, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 878, 14))
               End If
               If i2 = 5 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 892, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 942, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 957, 14))
               End If
               If i2 = 6 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 971, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1021, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1036, 14))
               End If
               If i2 = 7 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1050, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1100, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1115, 14))
               End If
               If i2 = 8 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1129, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1179, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1194, 14))
               End If
               If i2 = 9 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1208, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1258, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1273, 14))
               End If
               If i2 = 10 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1287, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1337, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1352, 14))
               End If
               If i2 = 11 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1366, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1416, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1431, 14))
               End If
               If i2 = 12 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1445, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1495, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1510, 14))
               End If
               If i2 = 13 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1524, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1574, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1589, 14))
               End If
               If i2 = 14 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1603, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1653, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1668, 14))
               End If
               If i2 = 15 Then
                  Reg.DESCRICAO(i2) = Trim(Mid(linha, 1682, 50))
                  Reg.FAIXA(i2) = Trim(Mid(linha, 1732, 15))
                  Reg.VALOR(i2) = Trim(Mid(linha, 1747, 14))
               End If
           Next
           
           Reg.TOTAL = Trim(Mid(linha, 1776, 14))
           
           For i3 = 1 To 3
                If i3 = 1 Then
                   Reg.MENSAGEM(i3) = Trim(Mid(linha, 1790, 100))
                End If
                If i3 = 2 Then
                   Reg.MENSAGEM(i3) = Trim(Mid(linha, 1890, 100))
                End If
                If i3 = 3 Then
                   Reg.MENSAGEM(i3) = Trim(Mid(linha, 1990, 100))
                End If
           Next
           
           Reg.GERENCIA_REGIONAL = Trim(Mid(linha, 2090, 25))
           Reg.MES_ANO = Trim(Mid(linha, 2115, 7))
           Reg.TURBIDEZ = Trim(Mid(linha, 2122, 9))
           Reg.CLORO = Trim(Mid(linha, 2131, 9))
           Reg.GRUPO = Trim(Mid(linha, 2140, 2))
           'Reg.GRUPO_EMPRESA = Trim(Mid(linha, 2136, 3))
           Reg.FIRMA = Trim(Mid(linha, 2142, 2))
           Reg.SEQ = Trim(Mid(linha, 2144, 7))
           Reg.LINHA_DIGITAVEL = Trim(Mid(linha, 2151, 59))
           Reg.LOCAL_PAGAMENTO = Trim(Mid(linha, 2210, 45))
           Reg.VENCIMENTO_FICHA = Trim(Mid(linha, 2255, 20))
           Reg.CEDENTE = Trim(Mid(linha, 2275, 50))
           Reg.AGENCIA_COD_CEDENTE = Trim(Mid(linha, 2325, 13))
           Reg.DATA_DOCUMENTO = Trim(Mid(linha, 2338, 10))
           Reg.N_DOCUMENTO = Trim(Mid(linha, 2348, 10))
           Reg.ESPECIE_DOC = Trim(Mid(linha, 2358, 3))
           Reg.ACEITE = Trim(Mid(linha, 2361, 1))
           Reg.DATA_PROCESSAMENTO = Trim(Mid(linha, 2362, 10))
           Reg.NOSSO_NUMERO = Trim(Mid(linha, 2372, 19))
           Reg.CARTEIRA = Trim(Mid(linha, 2391, 2))
           Reg.VALOR_TOTAL = Trim(Mid(linha, 2393, 14))
           Reg.Linha1_SACADO = Trim(Mid(linha, 2407, 55))
           Reg.Linha2_SACADO = Trim(Mid(linha, 2462, 73))
           Reg.BAR_CODE = Trim(Mid(linha, 2535, 112))
           
           If Len(Trim(Reg.CLIENTE)) > 0 Then
              MontaExtrato
              LimpaReg
              
              totRegs = totRegs + 1
              frmPrincipal.txQtdReg.Text = totRegs
           End If

           If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
              frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

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
        frmPrincipal.ProgressBarProcess.Value = 0.0001

   Exit Sub
'--------------------------------------------------------------------------------
Erro:
   Close #outFN
   Close #inFN
   
   frmPrincipal.MousePointer = vbDefault
   frmPrincipal.ProgressBarProcess.Value = 0.0001
   
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
        Open App.Path & "\Conta_Boleto_A4_Int.ps" For Input As #200
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
'        Print #outFN, "/ImageData2"
'        Print #outFN, "currentfile"
'        Print #outFN, "<< /Filter /SubFileDecode"
'        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
'        Print #outFN, ">> /ReusableStreamDecode filter"
'        Open App.Path & "\Conta_Boleto_A4_Ext.ps" For Input As #200
'        While Not EOF(200)
'            Line Input #200, wCabeca
'            Print #outFN, wCabeca
'        Wend
'        Close #200
'        Print #outFN, "*EOD*"
'        Print #outFN, "def"
'        Print #outFN, " "
'        Print #outFN, "/EPSForm2"
'        Print #outFN, "<< /FormType 1"
'        Print #outFN, "   /BBox [0 0 595 900]"
'        Print #outFN, "   /Matrix [ 1 0 0 1 0 0]"
'        Print #outFN, "   /PaintProc"
'        Print #outFN, "   { pop"
'        Print #outFN, "       /ostate save def"
'        Print #outFN, "         /showpage {} def"
'        Print #outFN, "         /setpagedevice /pop load def"
'        Print #outFN, "         ImageData2 0 setfileposition ImageData2 cvx exec"
'        Print #outFN, "       ostate restore"
'        Print #outFN, "   } bind"
'        Print #outFN, ">> def"
'        Print #outFN, "%% Fim da Segunda Imagem"
'        Print #outFN, " "
        '-------------------------------------------------------------------
        Print #outFN, "%%============================"
        Print #outFN, "%% Inicio dos Dados Variaveis"
        Print #outFN, "%%============================"
        Print #outFN, " "
        Print #outFN, "%Aplicação: CONTA COMPESA Boleto A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        Dim i2 As Integer
        Dim i3 As Integer
        Reg.TIPO_CONTA = ""
        Reg.ESCRITORIO = ""
        Reg.MATRICULA = ""
        Reg.CLIENTE = ""
        Reg.VENCIMENTO1 = ""
        Reg.VENCIMENTO_FICHA = ""
        Reg.INSCRICAO = ""
        Reg.ENDERECO = ""
        Reg.ENDERECO1 = ""
        Reg.FATURA = ""
        Reg.RESPONSAVEL = ""
        Reg.ENDERECO_ENTREGA = ""
        Reg.ENDERECO_ENTREGA1 = ""
        Reg.CEP = ""
        Reg.AGUA = ""
        Reg.ESGOTO = ""
        For i = 1 To 3
            Reg.ULTIMOS_CONSUMOS(i) = ""
        Next
        Reg.LEITURA_ANTERIOR = ""
        Reg.LEITURA_ATUAL = ""
        Reg.CONSUMO_M3 = ""
        Reg.CONSUMO_DIAS = ""
        Reg.CONSUMO_MEDIA = ""
        Reg.TYPE_CONSUMO1 = ""
        Reg.TYPE_CONSUMO2 = ""
        Reg.PERIODO1 = ""
        Reg.PERIODO2 = ""
        Reg.ECONOMIA = ""
        Reg.CONSUMO_POR_ECONOMIA = ""
        Reg.COD_AUXILIAR = ""
        For i2 = 1 To 15
            Reg.DESCRICAO(i2) = ""
            Reg.FAIXA(i2) = ""
            Reg.VALOR(i2) = ""
        Next
        Reg.TOTAL = ""
        For i3 = 1 To 3
            Reg.MENSAGEM(i3) = ""
        Next
        Reg.GERENCIA_REGIONAL = ""
        Reg.MES_ANO = ""
        Reg.TURBIDEZ = ""
        Reg.CLORO = ""
        Reg.SEQ = ""
        Reg.LINHA_DIGITAVEL = ""
        Reg.LOCAL_PAGAMENTO = ""
        Reg.DATA_DOCUMENTO = ""
        Reg.N_DOCUMENTO = ""
        Reg.ESPECIE_DOC = ""
        Reg.ACEITE = ""
        Reg.DATA_PROCESSAMENTO = ""
        Reg.NOSSO_NUMERO = ""
        Reg.CEDENTE = ""
        Reg.VENCIMENTO_BOLETO = ""
        Reg.AGENCIA_COD_CEDENTE = ""
        Reg.CARTEIRA = ""
        Reg.ESPECIE_MOEDA = ""
        Reg.BAR_CODE = ""
        Reg.NAO_RECEBER = ""
        Reg.GRUPO = ""
        Reg.FIRMA = ""
        Reg.VALOR_TOTAL = ""
        Reg.Linha1_SACADO = ""
        Reg.Linha2_SACADO = ""
End Sub
