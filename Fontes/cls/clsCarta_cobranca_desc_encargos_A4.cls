VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarta_cobranca_desc_enc_A4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap Consultoria
'Data: 27 de Novembro de 2008
'Contatos - Email: atendimento@sistemap.com.br
'           Fone: (81) 9979.1972 / 9832.6084
'Aplica��o: Aviso de Cobran�a Compesa A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Aviso de Cobran�a Compesa A4"

'Vari�veis relacionadas com os Arquivos de Entrada e Sa�da
Private inFN As Integer, outFN As Integer
'Pastas que cont�m os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que est�o sendo processados
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
        FATURA(24) As String
        VENCIMENTO(24) As String
        VALOR(24) As String
        ENCARGOS(24) As String
        VALOR_ATUALIZADO(24) As String
        TOTAL_VALOR As String
        TOTAL_ENCARGOS As String
        TOTAL_VALOR_ATUALIZADO As String
        FATURAS_EM_ATRASO As String
        VALOR_HISTORICO As String
        RESUMO_ENCARGOS As String
        DEBITO_VENCIDO As String
        VALOR_FICHA As String
        TOTAL_DOS_DESCONTOS As String
        TOTAL_SEM_DESCONTOS As String
End Type
'Vari�veis gerais para os tipos
Private Reg As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Integer
        Dim posY As Long
        Dim posX As Long
        
        wRec = wRec + 1
        posX = 340
        posY = 290
        
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
        posY = posY + 115
'        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
'        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO_BOLETO & ") show"
        
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
        posY = posY + 325
        For i = 1 To 24
            If i >= 1 And i <= 12 Then
               Print #outFN, posX - 210 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FATURA(i))) & Reg.FATURA(i) & ") show"
               Print #outFN, posX + 80 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO(i) & ") show"
               Print #outFN, posX + 180 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR(i))) & Reg.VALOR(i) & ") show"
               Print #outFN, posX + 360 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.ENCARGOS(i))) & Reg.ENCARGOS(i) & ") show"
               Print #outFN, posX + 580 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_ATUALIZADO(i))) & Reg.VALOR_ATUALIZADO(i) & ") show"
            Else
               If i = 13 Then posY = 847
               Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.FATURA(i))) & Reg.FATURA(i) & ") show"
               Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO(i) & ") show"
               Print #outFN, posX + 1180 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR(i))) & Reg.VALOR(i) & ") show"
               Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.ENCARGOS(i))) & Reg.ENCARGOS(i) & ") show"
               Print #outFN, posX + 1590 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.VALOR_ATUALIZADO(i))) & Reg.VALOR_ATUALIZADO(i) & ") show"
            End If

            posY = posY + 31
        Next
        posY = 1260
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8 scalefont setfont"
        If Len(Trim(Reg.FATURA(13))) > 0 Then
'           Print #outFN, posX + 920 & " pontos -" & posY & " pontos moveto (" & "Total" & ") show"
           Print #outFN, posX + 1150 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_VALOR)) & Reg.TOTAL_VALOR & ") show"
           Print #outFN, posX + 1345 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_ENCARGOS)) & Reg.TOTAL_ENCARGOS & ") show"
           Print #outFN, posX + 1555 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_SEM_DESCONTOS)) & Reg.TOTAL_SEM_DESCONTOS & ") show"
        Else
'           Print #outFN, posX - 90 & " pontos -" & posY & " pontos moveto (" & "Total" & ") show"
           Print #outFN, posX + 145 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_VALOR)) & Reg.TOTAL_VALOR & ") show"
           Print #outFN, posX + 330 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_ENCARGOS)) & Reg.TOTAL_ENCARGOS & ") show"
           Print #outFN, posX + 545 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_SEM_DESCONTOS)) & Reg.TOTAL_SEM_DESCONTOS & ") show"
        End If
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 10 scalefont setfont"
        posY = posY + 233
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.FATURAS_EM_ATRASO)) & Reg.FATURAS_EM_ATRASO & ") show"
        posY = posY + 60
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_VALOR)) & Reg.TOTAL_VALOR & ") show"
        posY = posY + 60
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_ENCARGOS)) & Reg.TOTAL_ENCARGOS & ") show"
        posY = posY + 70
        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_VALOR_ATUALIZADO)) & Reg.TOTAL_SEM_DESCONTOS & ") show"
        posY = posY + 70
'        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_DOS_DESCONTOS)) & Reg.TOTAL_DOS_DESCONTOS & ") show"
'        posY = posY + 60
'        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Space(18 - Len(Reg.TOTAL_SEM_DESCONTOS)) & Reg.TOTAL_VALOR_ATUALIZADO & ") show"
        
        posY = posY + 1155
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 515 & " pontos -" & posY + 22 & " pontos moveto (" & Reg.LINHA_DIGITAVEL & ") show"
        posY = posY + 195
        Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
        Print #outFN, posX + 300 & " pontos -" & posY + 10 & " pontos moveto (" & Reg.BAR_CODE & ") show"
        posY = posY + 125
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.ORDEM & ") show"
        Print #outFN, posX + 1650 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 120
        Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 690 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO_BOLETO & ") show"
        Print #outFN, posX + 1470 & " pontos -" & posY & " pontos moveto (" & Reg.TOTAL_VALOR_ATUALIZADO & ") show"
        Print #outFN, "showpage"
        'Fim da Parte Interna
        '**********Inicio Parte Externa************************************************
        posX = 810
        posY = 620
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

        'Gravar Cabe�alho
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
           Reg.TOTAL_DOS_DESCONTOS = Trim(Mid(linha, 2087, 14))
           Reg.TOTAL_SEM_DESCONTOS = Trim(Mid(linha, 2073, 14))

           For i1 = 1 To 24
               If i1 = 1 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 502, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 513, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 523, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 537, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 551, 14))
               End If
               If i1 = 2 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 565, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 576, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 586, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 600, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 614, 14))
               End If
               If i1 = 3 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 628, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 639, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 649, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 663, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 677, 14))
               End If
               If i1 = 4 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 691, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 702, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 712, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 726, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 740, 14))
               End If
               If i1 = 5 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 754, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 765, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 775, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 789, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 803, 14))
               End If
               If i1 = 6 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 817, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 828, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 838, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 852, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 866, 14))
               End If
               If i1 = 7 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 880, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 891, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 901, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 915, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 929, 14))
               End If
               If i1 = 8 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 943, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 954, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 964, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 978, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 992, 14))
               End If
               If i1 = 9 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1006, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1017, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1027, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1041, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1055, 14))
               End If
               If i1 = 10 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1069, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1080, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1090, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1104, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1118, 14))
               End If
               If i1 = 11 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1132, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1143, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1153, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1167, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1181, 14))
               End If
               If i1 = 12 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1195, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1206, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1216, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1231, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1245, 14))
               End If
               If i1 = 13 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1258, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1269, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1279, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1293, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1307, 14))
               End If
               If i1 = 14 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1321, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1332, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1342, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1356, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1370, 14))
               End If
               If i1 = 15 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1384, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1395, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1405, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1419, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1433, 14))
               End If
               If i1 = 16 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1447, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1458, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1468, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1482, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1496, 14))
               End If
               If i1 = 17 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1510, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1521, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1531, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1545, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1559, 14))
               End If
               If i1 = 18 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1573, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1584, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1594, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1608, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1622, 14))
               End If
               If i1 = 19 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1636, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1647, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1657, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1671, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1685, 14))
               End If
               If i1 = 20 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1699, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1710, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1720, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1734, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1748, 14))
               End If
               If i1 = 21 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1762, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1773, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1783, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1797, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1811, 14))
               End If
               If i1 = 22 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1825, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1836, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1846, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1860, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1874, 14))
               End If
               If i1 = 23 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1888, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1899, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1909, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1923, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 1937, 14))
               End If
               If i1 = 24 Then
                  Reg.FATURA(i1) = Trim(Mid(linha, 1951, 11))
                  Reg.VENCIMENTO(i1) = Trim(Mid(linha, 1962, 10))
                  Reg.VALOR(i1) = Trim(Mid(linha, 1972, 14))
                  Reg.ENCARGOS(i1) = Trim(Mid(linha, 1986, 14))
                  Reg.VALOR_ATUALIZADO(i1) = Trim(Mid(linha, 2000, 14))
               End If
           Next
           Reg.FATURAS_EM_ATRASO = Trim(Mid(linha, 2014, 3))
           Reg.VALOR_FICHA = Trim(Mid(linha, 2031, 14))
           Reg.TOTAL_VALOR = Trim(Mid(linha, 2017, 14))
           Reg.TOTAL_ENCARGOS = Trim(Mid(linha, 2045, 14))
           Reg.TOTAL_VALOR_ATUALIZADO = Trim(Mid(linha, 2059, 14))

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
   
   MsgBox "Descri��o do Erro: " & Err.Description & vbCrLf & _
          "N�mero do Erro:    " & Err.Number & vbCrLf & _
          "ERRO durante o processamento do arquivo " & procArq & "." & vbCrLf & _
          "Caso o arquivo " & gerArq & vbCrLf & _
          "ainda exista, favor APAG�-LO E FAZER TODO O PROCESSO DO IN�CIO.", vbCritical
   
   If Manip_Arq(EXISTEARQ, gerArq) Then
      If Not Manip_Arq(DELETAR, gerArq) Then
         MsgBox "O arquivo " & gerArq & " n�o pode ser exclu�do." & _
         vbCrLf & "Favor apag�-lo e reiniciar o processamento.", vbCritical
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
        'Cabe�alho
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
        Open App.Path & "\carta_cobranca_desc_encargos.ps" For Input As #200
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
        Open App.Path & "\Aviso_De_Cobranca_Compesa_A4_Ext.ps" For Input As #200
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
        Print #outFN, "%Aplica��o: Aviso de Cobran�a Fim do Ano Compesa A4"
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
        Reg.TOTAL_VALOR = ""
        Reg.TOTAL_ENCARGOS = ""
        Reg.TOTAL_VALOR_ATUALIZADO = ""
        Reg.FATURAS_EM_ATRASO = ""
        Reg.VALOR_HISTORICO = ""
        Reg.RESUMO_ENCARGOS = ""
        Reg.DEBITO_VENCIDO = ""
        Reg.VALOR_FICHA = ""
        For i = 1 To 24
            Reg.FATURA(i) = ""
            Reg.VENCIMENTO(i) = ""
            Reg.VALOR(i) = ""
            Reg.ENCARGOS(i) = ""
            Reg.VALOR_ATUALIZADO(i) = ""
        Next
End Sub
