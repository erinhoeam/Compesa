VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConta_Compesa_A5"
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
'Aplicação: Conta Compesa A4/A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double
Dim wQtdFaturas(2) As Long
Dim wIndiceFaturas As Long

Const AppName = "COMPESA", AppDesc = "Conta Compesa A4/A5"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        MENSAGEM  As String  '  1   50  50
        Nome_Localidade  As String  '   51  80  30
        Matrícula_Imovel  As String  '  81  90  10
        MES_ANO  As String  '  de Referência e dígito verificador da Conta 91  99  9
        NOME_CONTA   As String  '   100 149 50
        ENDERECO_IMOVEL  As String  '   150 199 50
        Endereco_Imovel2  As String  '  - Segunda Linha  200 249 50
        INSCRICAO   As String  '   250 269 20
        Grupo_Faturamento  As String  '     270 271 2
        Codigo_debito_automatico   As String  '    272 281 10
        Codigo_cliente_responsável  As String  '    282 289 8
        Endereco_correspondencia_responsavel  As String  '   290 339 50
        Endereco_correspondencia_responsavel2   As String  '   - Segunda Linha  340 389 50
        situacao_ligacao_agua   As String  '    390 409 20
        situacao_ligacao_esgoto   As String  '  410 429 20
        Economias_Residencial  As String  '   430 432 3
        Economias_Comercial  As String  '     433 435 3
        Economias_Industrial As String  '     436 438 3
        Economias_Publico  As String  '   439 441 3
        Hidrometro   As String  '  442 451 10
        Data_Leitura_Anterior  As String  '     452 461 10
        Data_Leitura_Atual As String  '     462 471 10
        Tipo_Consumo  As String  '    472 512 41
        Leitura_Anterior_Agua  As String  '  513 519 7
        Constante_Consumo_Faturado_Agua    As String  '  520 527 8
        Consumo_Faturado_Agua  As String  '  528 532 5
        Leitura_Atual_Agua  As String  '     533 539 7
        Anormalidade_Leitura_Agua  As String  '  540 564 25
        Leitura_Faturada_Agua  As String  '  565 571 7
        Anormalidade_Consumo_Agua  As String  '  572 596 25
        Leitura_Anterior_Esgoto  As String  '    597 603 7
        Constante_Consumo_Faturado_Esgoto  As String  '  604 611 8
        Consumo_Faturado_Esgoto  As String  '    612 616 5
        Leitura_Atual_Esgoto  As String  '   617 623 7
        Anormalidade_Leitura_Esgoto   As String  '   624 648 25
        Leitura_Fat_Esgoto   As String  '   649 655 7
        Anormalidade_Consumo_Esgoto  As String  '    656 680 25
        Consumo_Mes1  As String  '    681 699 19
        Consumo_Mes2  As String  '   700 718 19
        Consumo_Mes3  As String  '   719 737 19
        Consumo_Mes4  As String  '    738 756 19
        Consumo_Mes5  As String  '   757 775 19
        Consumo_Mes6  As String  '   776 794 19
        A1  As String  '    681 699 19
        A2  As String  '    681 699 19
        B1  As String  '    681 699 19
        B2  As String  '    681 699 19
        C1  As String  '    681 699 19
        C2  As String  '    681 699 19
        D1  As String  '    681 699 19
        D2  As String  '    681 699 19
        E1  As String  '    681 699 19
        E2  As String  '    681 699 19
        F1  As String  '    681 699 19
        F2  As String  '    681 699 19
        M1  As String  '    681 699 19
        M2  As String  '    681 699 19
        
        MEDIA_CONSUMO As String  '     795 813 19
        Turbidez_Exigida   As String  '    814 818 5
        Turbidez_Realizada  As String  '   819 823 5
        Turbidez_Conforme    As String  '  824 828 5
        Cor_Aparente_Exigida   As String  '    829 833 5
        Cor_Aparente_Realizada  As String  '   834 838 5
        Cor_Aparente_Conforme    As String  '  839 843 5
        Cloro_Residual_Livre_Exigida  As String  '     844 848 5
        Cloro_Residual_Livre_Realizada  As String  '   849 853 5
        Cloro_Residual_Livre_Conforme  As String  '    854 858 5
        Coliformes_Totais_Exigida    As String  '  859 863 5
        Coliformes_Totais_Realizada  As String  '  864 868 5
        Coliformes_Totais_Conforme   As String  '  869 873 5
        E_Coli_Exigida  As String  '  874 878 5
        E_Coli_Realizada    As String  '  879 883 5
        E_Coli_Conforme     As String  '  884 888 5
        Observacao1  As String  '  889 950 62
        Observacao2  As String  '  951 1070    120
        Observacao3  As String  '  1071    1170    100
        Servicos1   As String  '  1171    1220    50
        Consumo_Faixa1  As String  '  1221    1235    15
        Valor1  As String  '  1236    1249    14
        Servicos2   As String  '  1250    1299    50
        Consumo_Faixa2  As String  '  1300    1314    15
        Valor2  As String  '  1315    1328    14
        Servicos3  As String  '   1329    1378    50
        Consumo_Faixa3  As String  '   1379    1393    15
        Valor3  As String  '  1394    1407    14
        Servicos4   As String  '  1408    1457    50
        Consumo_Faixa4  As String  '  1458    1472    15
        Valor4  As String  '  1473    1486    14
        Servicos5   As String  '  1487    1536    50
        Consumo_Faixa5  As String  '  1537    1551    15
        Valor5  As String  '  1552    1565    14
        Servicos6   As String  '  1566    1615    50
        Consumo_Faixa6  As String  '  1616    1630    15
        Valor6  As String  '  1631    1644    14
        Servicos7   As String  '  1645    1694    50
        Consumo_Faixa7  As String  '  1695    1709    15
        Valor7  As String  '  1710    1723    14
        Servicos8   As String  '  1724    1773    50
        Consumo_Faixa8  As String  '  1774    1788    15
        Valor8  As String  '  1789    1802    14
        Servicos9   As String  '  1803    1852    50
        Consumo_Faixa9  As String  '  1853    1867    15
        Valor9  As String  '  1868    1881    14
        Servicos10  As String  '  1882    1931    50
        Consumo_Faixa10   As String  '  1932    1946    15
        Valor10   As String  '   1947    1960    14
        Servicos11  As String  ' 1961    2010    50
        Consumo11  As String  '    2011    2025    15
        Valor11     As String  ' 2026    2039    14
        Servicos12  As String  ' 2040    2089    50
        Consumo12   As String  '   2090    2104    15
        Valor12   As String  '   2105    2118    14
        Servicos13  As String  ' 2119    2168    50
        Consumo_Faixa13  As String  '    2169    2183    15
        Valor13   As String  '   2184    2197    14
        Servicos14  As String  ' 2198    2247    50
        Consumo_Faixa14   As String  '   2248    2262    15
        Valor14   As String  '   2263    2276    14
        Servicos15  As String  ' 2277    2326    50
        ConsumoFaixa15     As String  ' 2327    2341    15
        Valor15   As String  '   2342    2355    14
        Mensagem_rateio_consumo_fixo_esgoto      As String  ' 2356    2386    31
        PIS_Base_Calculo  As String  ' 2387    2400    14
        PIS_Percentual   As String  ' 2401    2405    5
        PIS_Valor_Imposto    As String  ' 2406    2419    14
        COFINS_Base_Calculo   As String  ' 2420    2433    14
        COFINS_Percentual    As String  ' 2434    2438    5
        COFINS_Valor_Imposto   As String  '   2439    2452    14
        Data_Vencimento   As String  ' 2453    2462    10
        Valor_Conta   As String  ' 2463    2476    14
        Mensagem_conta1   As String  '  2477    2576    100
        Mensagem_conta2   As String  '  2577    2676    100
        Mensagem_conta3   As String  ' 2677    2776    100
        SEQUENCIAL   As String  ' 3000    3010   10
        Código_Banco  As String  ' 2784    2788    5
        Representacao_numerica_codigo_barras   As String  ' 2789    2842    54
        Local_Pagamento   As String  ' 2843    2887    45
        VENCIMENTO   As String  ' 2888    2907    20
        Cedente  As String  ' 2908    2957    50
        Agencia_Cedente    As String  ' 2958    2970    13
        Data_Documento     As String  ' 2971    2980    10
        Numero_Documento       As String  ' 2981    2990    10
        Especie_Documento      As String  ' 2991    2993    3
        Aceite    As String  ' 2994    2994    1
        Data_Processamento     As String  ' 2995    3004    10
        Nosso_Numero      As String  ' 3005    3023    19
        Carteira      As String  ' 3024    3025    2
        Valor_Documento    As String  ' 3026    3039    14
        Sacado_linha_1a    As String  ' 3040    3089    50
        Sacado_linha1b   As String  '  3090    3105    16
        Sacado_linha1c    As String  ' 3106    3115    10
        Sacado_linha1d    As String  ' 3116    3129    14
        Sacado_linha1e   As String  '  3130    3138    9
        Sacado_linha2a    As String  ' 3139    3188    50
        Sacado_linha3a    As String  ' 3189    3238    50
        Codigo_Barras      As String  ' 3239    3350    112
        SEQ      As String  ' 3239    3350    112
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim i2 As Long
        Dim i3 As Long
        Dim i4 As Long
        Dim posY As Long
        Dim posX As Long
        
        If Len(Trim(Reg(1).NOME_CONTA)) Then wRec = wRec + 1
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        For i = 1 To 2
            
            If i = 1 Then
              'Print #outFN, "save EPSForm1 execform restore"
              'Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
            End If
            
           'Capturar o Maior Consumo
            Dim wMaior As Long
            Dim wDivisor As Long
            wMaior = 0

            If Val(Reg(i).A1) > wMaior Then wMaior = Val(Reg(i).A1)
            If Val(Reg(i).A2) > wMaior Then wMaior = Val(Reg(i).A2)
            If Val(Reg(i).B1) > wMaior Then wMaior = Val(Reg(i).B1)
            If Val(Reg(i).B2) > wMaior Then wMaior = Val(Reg(i).B2)
            If Val(Reg(i).C1) > wMaior Then wMaior = Val(Reg(i).C1)
            If Val(Reg(i).C2) > wMaior Then wMaior = Val(Reg(i).C2)
            If Val(Reg(i).D1) > wMaior Then wMaior = Val(Reg(i).D1)
            If Val(Reg(i).D2) > wMaior Then wMaior = Val(Reg(i).D2)
            If Val(Reg(i).E1) > wMaior Then wMaior = Val(Reg(i).E1)
            If Val(Reg(i).E2) > wMaior Then wMaior = Val(Reg(i).E2)
            If Val(Reg(i).F1) > wMaior Then wMaior = Val(Reg(i).F1)
            If Val(Reg(i).F2) > wMaior Then wMaior = Val(Reg(i).F2)
            
            'Testes
             wDivisor = 1
             If wMaior >= 1 And wMaior < 50 Then wDivisor = 1
             If wMaior >= 50 And wMaior < 100 Then wDivisor = 2
             If wMaior >= 100 And wMaior < 200 Then wDivisor = 14
             If wMaior >= 200 And wMaior < 300 Then wDivisor = 9
             If wMaior >= 300 And wMaior < 400 Then wDivisor = 10
             If wMaior >= 400 And wMaior < 500 Then wDivisor = 11
             If wMaior >= 500 And wMaior < 600 Then wDivisor = 12
             If wMaior >= 600 And wMaior < 700 Then wDivisor = 12.7
             If wMaior >= 700 And wMaior < 800 Then wDivisor = 13.7
             If wMaior >= 800 And wMaior < 900 Then wDivisor = 14.7
             If wMaior >= 900 And wMaior < 1000 Then wDivisor = 15.5
             If wMaior >= 1000 And wMaior < 1100 Then wDivisor = 16
             If wMaior >= 1100 And wMaior < 1200 Then wDivisor = 17.5
             If wMaior >= 1200 And wMaior < 1300 Then wDivisor = 22
             If wMaior >= 1300 And wMaior < 1400 Then wDivisor = 23
             If wMaior >= 1400 And wMaior < 1500 Then wDivisor = 24
             If wMaior >= 1500 And wMaior < 1600 Then wDivisor = 25
             If wMaior >= 1600 And wMaior < 1700 Then wDivisor = 27
             If wMaior >= 1700 And wMaior < 1800 Then wDivisor = 29
             If wMaior >= 1800 And wMaior < 1900 Then wDivisor = 31
             If wMaior >= 1900 And wMaior < 2000 Then wDivisor = 33
             If wMaior >= 2000 And wMaior < 2100 Then wDivisor = 36
             If wMaior >= 2100 And wMaior < 2200 Then wDivisor = 45
             If wMaior >= 2200 And wMaior < 2300 Then wDivisor = 46
             If wMaior >= 2300 And wMaior < 2400 Then wDivisor = 47
             If wMaior >= 2400 And wMaior < 2500 Then wDivisor = 48
             If wMaior >= 2500 And wMaior < 2600 Then wDivisor = 49
             If wMaior >= 2600 And wMaior < 2700 Then wDivisor = 50
             If wMaior >= 2700 And wMaior < 2800 Then wDivisor = 51
             If wMaior >= 2800 And wMaior < 2900 Then wDivisor = 52
             If wMaior >= 2900 And wMaior < 3000 Then wDivisor = 53
             If wMaior >= 3000 And wMaior < 3100 Then wDivisor = 54
             If wMaior >= 3100 And wMaior < 3200 Then wDivisor = 55
             If wMaior >= 3200 And wMaior < 3300 Then wDivisor = 56
             If wMaior >= 3300 And wMaior < 3400 Then wDivisor = 57
             If wMaior >= 3500 And wMaior < 3500 Then wDivisor = 58
             If wMaior >= 3500 Then wDivisor = 59
           '-------------------------------------------------------
           If i = 1 Then
            If Val(Trim(Reg(i).A1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "159 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).A1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).A1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i).A2)) > 0 And wMaior > 0 Then
               Print #outFN, "162 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).A2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).A2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i).B1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Segunda Barra"
               Print #outFN, "newpath"
               Print #outFN, "167 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).B1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).B1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i).B2)) > 0 And wMaior > 0 Then
               Print #outFN, "170 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).B2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).B2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i).C1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Terceira Barra"
               Print #outFN, "newpath"
               Print #outFN, "177 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).C1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).C1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i).C2)) > 0 And wMaior > 0 Then
               Print #outFN, "180 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).C2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).C2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i).D1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Quarta Barra"
               Print #outFN, "newpath"
               Print #outFN, "186 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).D1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).D1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i).D2)) > 0 And wMaior > 0 Then
               Print #outFN, "189 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).D2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).D2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i).E1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Quinta Barra"
               Print #outFN, "newpath"
               Print #outFN, "194 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).E1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).E1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i).E2)) > 0 And wMaior > 0 Then
               Print #outFN, "197 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).E2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).E2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i).F1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Sexta Barra"
               Print #outFN, "newpath"
               Print #outFN, "203 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).F1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).F1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i).F2)) > 0 And wMaior > 0 Then
               Print #outFN, "206 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).F2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).F2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i).M1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Media Barra"
               Print #outFN, "newpath"
               Print #outFN, "212 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).M1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).M1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i).M2)) > 0 And wMaior > 0 Then
               Print #outFN, "215 85 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i).M2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i).M2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
           ' Ficha do Lado Direito
           '-------------------------------------------------------
            wMaior = 0

            If Val(Reg(2).A1) > wMaior Then wMaior = Val(Reg(2).A1)
            If Val(Reg(2).A2) > wMaior Then wMaior = Val(Reg(2).A2)
            If Val(Reg(2).B1) > wMaior Then wMaior = Val(Reg(2).B1)
            If Val(Reg(2).B2) > wMaior Then wMaior = Val(Reg(2).B2)
            If Val(Reg(2).C1) > wMaior Then wMaior = Val(Reg(2).C1)
            If Val(Reg(2).C2) > wMaior Then wMaior = Val(Reg(2).C2)
            If Val(Reg(2).D1) > wMaior Then wMaior = Val(Reg(2).D1)
            If Val(Reg(2).D2) > wMaior Then wMaior = Val(Reg(2).D2)
            If Val(Reg(2).E1) > wMaior Then wMaior = Val(Reg(2).E1)
            If Val(Reg(2).E2) > wMaior Then wMaior = Val(Reg(2).E2)
            If Val(Reg(2).F1) > wMaior Then wMaior = Val(Reg(2).F1)
            If Val(Reg(2).F2) > wMaior Then wMaior = Val(Reg(2).F2)
            
            'Testes
             wDivisor = 1
             If wMaior >= 1 And wMaior < 50 Then wDivisor = 1
             If wMaior >= 50 And wMaior < 100 Then wDivisor = 2
             If wMaior >= 100 And wMaior < 200 Then wDivisor = 14
             If wMaior >= 200 And wMaior < 300 Then wDivisor = 9
             If wMaior >= 300 And wMaior < 400 Then wDivisor = 10
             If wMaior >= 400 And wMaior < 500 Then wDivisor = 11
             If wMaior >= 500 And wMaior < 600 Then wDivisor = 12
             If wMaior >= 600 And wMaior < 700 Then wDivisor = 12.7
             If wMaior >= 700 And wMaior < 800 Then wDivisor = 13.7
             If wMaior >= 800 And wMaior < 900 Then wDivisor = 14.7
             If wMaior >= 900 And wMaior < 1000 Then wDivisor = 15.5
             If wMaior >= 1000 And wMaior < 1100 Then wDivisor = 16
             If wMaior >= 1100 And wMaior < 1200 Then wDivisor = 17.5
             If wMaior >= 1200 And wMaior < 1300 Then wDivisor = 22
             If wMaior >= 1300 And wMaior < 1400 Then wDivisor = 23
             If wMaior >= 1400 And wMaior < 1500 Then wDivisor = 24
             If wMaior >= 1500 And wMaior < 1600 Then wDivisor = 25
             If wMaior >= 1600 And wMaior < 1700 Then wDivisor = 27
             If wMaior >= 1700 And wMaior < 1800 Then wDivisor = 29
             If wMaior >= 1800 And wMaior < 1900 Then wDivisor = 31
             If wMaior >= 1900 And wMaior < 2000 Then wDivisor = 33
             If wMaior >= 2000 And wMaior < 2100 Then wDivisor = 36
             If wMaior >= 2100 And wMaior < 2200 Then wDivisor = 45
             If wMaior >= 2200 And wMaior < 2300 Then wDivisor = 46
             If wMaior >= 2300 And wMaior < 2400 Then wDivisor = 47
             If wMaior >= 2400 And wMaior < 2500 Then wDivisor = 48
             If wMaior >= 2500 And wMaior < 2600 Then wDivisor = 49
             If wMaior >= 2600 And wMaior < 2700 Then wDivisor = 50
             If wMaior >= 2700 And wMaior < 2800 Then wDivisor = 51
             If wMaior >= 2800 And wMaior < 2900 Then wDivisor = 52
             If wMaior >= 2900 And wMaior < 3000 Then wDivisor = 53
             If wMaior >= 3000 And wMaior < 3100 Then wDivisor = 54
             If wMaior >= 3100 And wMaior < 3200 Then wDivisor = 55
             If wMaior >= 3200 And wMaior < 3300 Then wDivisor = 56
             If wMaior >= 3300 And wMaior < 3400 Then wDivisor = 57
             If wMaior >= 3500 And wMaior < 3500 Then wDivisor = 58
             If wMaior >= 3500 Then wDivisor = 59
           '-------------------------------------------------------
            If Val(Trim(Reg(i + 1).A1)) And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "159 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).A1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).A1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i + 1).A2)) > 0 And wMaior > 0 Then
               Print #outFN, "162 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).A2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).A2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i + 1).B1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "167 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).B1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).B1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i + 1).B2)) > 0 And wMaior > 0 Then
               Print #outFN, "170 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).B2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).B2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i + 1).C1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "177 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).C1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).C1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i + 1).C2)) > 0 And wMaior > 0 Then
               Print #outFN, "180 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).C2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).C2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i + 1).D1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "186 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).D1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).D1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i + 1).D2)) > 0 And wMaior > 0 Then
               Print #outFN, "189 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).D2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).D2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i + 1).E1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "194 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).E1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).E1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i + 1).E2)) > 0 And wMaior > 0 Then
               Print #outFN, "197 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).E2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).E2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i + 1).F1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "203 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).F1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).F1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i + 1).F2)) > 0 And wMaior > 0 Then
               Print #outFN, "206 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).F2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).F2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
           '-------------------------------------------------------
            If Val(Trim(Reg(i + 1).M1)) > 0 And wMaior > 0 Then
               Print #outFN, "%Primeira Barra"
               Print #outFN, "newpath"
               Print #outFN, "212 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).M1)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).M1 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            If Val(Trim(Reg(i + 1).M2)) > 0 And wMaior > 0 Then
               Print #outFN, "215 522 moveto"
               Print #outFN, "0 -3 rlineto"
               Print #outFN, "3 0 rlineto"
               Print #outFN, "0 " & CalcularBarra(wMaior, CLng(Reg(i + 1).M2)) & " rlineto"
               'Print #outFN, "0 " & (Int(Reg(i + 1).M2 / wDivisor)) & " rlineto"
               Print #outFN, "-3 0 rlineto"
               Print #outFN, "closepath"
               Print #outFN, "gsave"
               Print #outFN, "0.5 1 0.5 setrgbcolor"
               Print #outFN, "fill"
               Print #outFN, "grestore"
               Print #outFN, "0 0 0 setrgbcolor"
               Print #outFN, "1 setlinewidth"
               Print #outFN, "stroke"
            End If
            End If
           '-------------------------------------------------------
            Print #outFN, "0 0 0 setrgbcolor"
       
            If i = 1 Then
              If frmPrincipal.chkContaNResp.Value = 1 Then
                Print #outFN, "save EPSForm3 execform restore"
              End If
              If frmPrincipal.chkcontaN.Value = 1 Then
                Print #outFN, "save EPSForm1 execform restore"
              End If
              Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
            End If
            
            Select Case i
                Case 1
                  'posX = 155
                  'posY = 233
                  posX = 100
                  posY = 209
                Case 2
                  posX = 1920
                  posY = 209
            End Select
            
            'Rota
            Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
            Print #outFN, posX + 210 & " pontos -" & posY & " pontos moveto (" & Reg(i).Nosso_Numero & ") show"
            Print #outFN, posX + 630 & " pontos -" & posY + 7 & " pontos moveto (" & Reg(i).Nome_Localidade & ") show"
            posY = posY + 85
            'Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
            Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Matrícula_Imovel & ") show"
            Print #outFN, posX + 1320 & " pontos -" & posY & " pontos moveto (" & Reg(i).MES_ANO & ") show"
            posY = posY + 10
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 5.5 scalefont setfont"
            Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME_CONTA & ") show"
            posY = posY + 25
            Print #outFN, "/Helvetica-Condensed-AP findfont 5.5 scalefont setfont"
            Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Trim(Reg(i).ENDERECO_IMOVEL) & Trim(Reg(i).Endereco_Imovel2) & ") show"
    
            Print #outFN, posX + 150 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 770 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).Grupo_Faturamento & ") show"
            Print #outFN, posX + 1330 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).Codigo_debito_automatico & ") show"
        
            posY = posY + 77
            Print #outFN, "/Helvetica-Condensed-AP findfont 5.5 scalefont setfont"
            Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).Codigo_cliente_responsável & ") show"
            Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Trim(Reg(i).Endereco_correspondencia_responsavel) & Trim(Reg(i).Endereco_correspondencia_responsavel2) & ") show"
            posY = posY + 60
            Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg(i).situacao_ligacao_agua & ") show"
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg(i).situacao_ligacao_esgoto & ") show"
            
            Print #outFN, posX + 620 & " pontos -" & posY + 8 & " pontos moveto (" & Reg(i).Economias_Residencial & ") show"
            Print #outFN, posX + 850 & " pontos -" & posY + 8 & " pontos moveto (" & Reg(i).Economias_Comercial & ") show"
            Print #outFN, posX + 1080 & " pontos -" & posY + 8 & " pontos moveto (" & Reg(i).Economias_Industrial & ") show"
            Print #outFN, posX + 1330 & " pontos -" & posY + 8 & " pontos moveto (" & Reg(i).Economias_Publico & ") show"
            
            posY = posY + 55
            Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg(i).Hidrometro & ") show"
            Print #outFN, posX + 310 & " pontos -" & posY & " pontos moveto (" & Reg(i).Data_Leitura_Anterior & ") show"
            Print #outFN, posX + 680 & " pontos -" & posY & " pontos moveto (" & Reg(i).Data_Leitura_Atual & ") show"
            Print #outFN, posX + 990 & " pontos -" & posY & " pontos moveto (" & Reg(i).Tipo_Consumo & ") show"
            
           'Leituras
            Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
            posY = posY + 40
            Print #outFN, posX + 130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Leitura_Anterior_Agua & ") show"
            Print #outFN, posX + 900 & " pontos -" & posY & " pontos moveto (" & Reg(i).Leitura_Anterior_Esgoto & ") show"
            
            Print #outFN, posX + 480 & " pontos -" & posY & " pontos moveto (" & Reg(i).Consumo_Faturado_Agua & ") show"
            Print #outFN, posX + 1270 & " pontos -" & posY & " pontos moveto (" & Reg(i).Consumo_Faturado_Esgoto & ") show"
            
            posY = posY + 27
            Print #outFN, posX + 130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Leitura_Atual_Agua & ") show"
            Print #outFN, posX + 900 & " pontos -" & posY & " pontos moveto (" & Reg(i).Leitura_Atual_Esgoto & ") show"
            
            posY = posY + 27
            Print #outFN, posX + 130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Leitura_Faturada_Agua & ") show"
            Print #outFN, posX + 900 & " pontos -" & posY & " pontos moveto (" & Reg(i).Leitura_Fat_Esgoto & ") show"
            
            Print #outFN, "/Helvetica-Condensed-AP findfont 5.4 scalefont setfont"
            Print #outFN, posX + 350 & " pontos -" & posY - 4 & " pontos moveto (" & Reg(i).Anormalidade_Consumo_Agua & ") show"
           
           'Consumo
            Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
            posY = posY + 76
            Print #outFN, posX + 25 & " pontos -" & posY & " pontos moveto (" & Reg(i).Consumo_Mes1 & ") show"
            Print #outFN, posX + 25 & " pontos -" & posY + 35 & " pontos moveto (" & Reg(i).Consumo_Mes2 & ") show"
            Print #outFN, posX + 25 & " pontos -" & posY + 70 & " pontos moveto (" & Reg(i).Consumo_Mes3 & ") show"
            Print #outFN, posX + 25 & " pontos -" & posY + 105 & " pontos moveto (" & Reg(i).Consumo_Mes4 & ") show"
            Print #outFN, posX + 25 & " pontos -" & posY + 140 & " pontos moveto (" & Reg(i).Consumo_Mes5 & ") show"
            Print #outFN, posX + 25 & " pontos -" & posY + 175 & " pontos moveto (" & Reg(i).Consumo_Mes6 & ") show"
            Print #outFN, posX + 25 & " pontos -" & posY + 210 & " pontos moveto (" & Reg(i).MEDIA_CONSUMO & ") show"
            
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 4 scalefont setfont"
            posY = posY - 20
            Print #outFN, posX + 210 & " pontos -" & posY + 5 & " pontos moveto (A) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 20 & " pontos moveto (E) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 40 & " pontos moveto (A) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 55 & " pontos moveto (E) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 80 & " pontos moveto (A) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 95 & " pontos moveto (E) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 115 & " pontos moveto (A) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 130 & " pontos moveto (E) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 150 & " pontos moveto (A) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 165 & " pontos moveto (E) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 185 & " pontos moveto (A) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 200 & " pontos moveto (E) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 220 & " pontos moveto (A) show"
            Print #outFN, posX + 210 & " pontos -" & posY + 235 & " pontos moveto (E) show"
        
        
            'Turbidez
            Print #outFN, "/Helvetica-Condensed-AP findfont 5 scalefont setfont"
            posY = posY + 52
            Print #outFN, posX + 870 & " pontos -" & posY & " pontos moveto (" & Reg(i).Turbidez_Exigida & ") show"
            Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Turbidez_Realizada & ") show"
            Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg(i).Turbidez_Conforme & ") show"
            
            'Cor Aparente
            posY = posY + 26
            Print #outFN, posX + 870 & " pontos -" & posY & " pontos moveto (" & Reg(i).Cor_Aparente_Exigida & ") show"
            Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Cor_Aparente_Realizada & ") show"
            Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg(i).Cor_Aparente_Conforme & ") show"
            
            'Cloro Residual
            posY = posY + 26
            Print #outFN, posX + 870 & " pontos -" & posY & " pontos moveto (" & Reg(i).Cloro_Residual_Livre_Exigida & ") show"
            Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Cloro_Residual_Livre_Realizada & ") show"
            Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg(i).Cloro_Residual_Livre_Conforme & ") show"
            
            'Cloro Residual
            posY = posY + 26
            Print #outFN, posX + 870 & " pontos -" & posY & " pontos moveto (" & Reg(i).Coliformes_Totais_Exigida & ") show"
            Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg(i).Coliformes_Totais_Realizada & ") show"
            Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg(i).Coliformes_Totais_Conforme & ") show"
            
            'E Coli
            posY = posY + 26
            Print #outFN, posX + 870 & " pontos -" & posY & " pontos moveto (" & Reg(i).E_Coli_Exigida & ") show"
            Print #outFN, posX + 1130 & " pontos -" & posY & " pontos moveto (" & Reg(i).E_Coli_Realizada & ") show"
            Print #outFN, posX + 1350 & " pontos -" & posY & " pontos moveto (" & Reg(i).E_Coli_Conforme & ") show"
                    
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 4 scalefont setfont"
            Print #outFN, posX + 630 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).Observacao1 & ") show"
            Print #outFN, posX + 487 & " pontos -" & posY + 52 & " pontos moveto (" & Reg(i).Observacao2 & ") show"
            Print #outFN, posX + 487 & " pontos -" & posY + 74 & " pontos moveto (" & Reg(i).Observacao3 & ") show"

           'DESCRIÇÃO DOS SERVIÇOS
            posY = posY + 170
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 6 scalefont setfont"
            
            Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos1 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa1)) & Reg(i).Consumo_Faixa1 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor1)) & Reg(i).Valor1 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos2 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa2)) & Reg(i).Consumo_Faixa2 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor2)) & Reg(i).Valor2 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos3 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa3)) & Reg(i).Consumo_Faixa3 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor3)) & Reg(i).Valor3 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos4 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa4)) & Reg(i).Consumo_Faixa4 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor4)) & Reg(i).Valor4 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos5 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa5)) & Reg(i).Consumo_Faixa5 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor5)) & Reg(i).Valor5 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos6 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa6)) & Reg(i).Consumo_Faixa6 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor6)) & Reg(i).Valor6 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos7 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa7)) & Reg(i).Consumo_Faixa7 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor7)) & Reg(i).Valor7 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos8 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa8)) & Reg(i).Consumo_Faixa8 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor8)) & Reg(i).Valor8 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos9 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa9)) & Reg(i).Consumo_Faixa9 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor9)) & Reg(i).Valor9 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos10 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa10)) & Reg(i).Consumo_Faixa10 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor10)) & Reg(i).Valor10 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos11 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo11)) & Reg(i).Consumo11 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor11)) & Reg(i).Valor11 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos12 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo12)) & Reg(i).Consumo12 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor12)) & Reg(i).Valor12 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos13 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa13)) & Reg(i).Consumo_Faixa13 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor13)) & Reg(i).Valor13 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos14 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).Consumo_Faixa14)) & Reg(i).Consumo_Faixa14 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor14)) & Reg(i).Valor14 & ") show"
            posY = posY + 25
            Print #outFN, posX + 10 & " pontos -" & posY & " pontos moveto (" & Reg(i).Servicos15 & ") show"
            Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).ConsumoFaixa15)) & Reg(i).ConsumoFaixa15 & ") show"
            Print #outFN, posX + 1090 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).Valor15)) & Reg(i).Valor15 & ") show"
            
            posY = posY + 50
            Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg(i).Mensagem_rateio_consumo_fixo_esgoto & ") show"
            
            posY = posY + 132
            Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 5.5 scalefont setfont"
            Print #outFN, posX + 270 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).PIS_Base_Calculo)) & Reg(i).PIS_Base_Calculo & ") show"
            Print #outFN, posX + 640 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).PIS_Percentual)) & Reg(i).PIS_Percentual & ") show"
            Print #outFN, posX + 880 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).PIS_Valor_Imposto)) & Reg(i).PIS_Valor_Imposto & ") show"
            
            posY = posY + 25
            Print #outFN, posX + 270 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).COFINS_Base_Calculo)) & Reg(i).COFINS_Base_Calculo & ") show"
            Print #outFN, posX + 640 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg(i).COFINS_Percentual)) & Reg(i).COFINS_Percentual & ") show"
            Print #outFN, posX + 880 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg(i).COFINS_Valor_Imposto)) & Reg(i).COFINS_Valor_Imposto & ") show"
            
            posY = posY + 63
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
            Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (" & Reg(i).Data_Vencimento & ") show"
            Print #outFN, posX + 1250 & " pontos -" & posY & " pontos moveto (" & Reg(i).Valor_Conta & ") show"
                
            posY = posY + 58
            Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
            Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg(i).Mensagem_conta1 & ") show"
            Print #outFN, posX + 30 & " pontos -" & posY + 30 & " pontos moveto (" & Reg(i).Mensagem_conta2 & ") show"
            Print #outFN, posX + 30 & " pontos -" & posY + 60 & " pontos moveto (" & Reg(i).Mensagem_conta3 & ") show"
            
            posY = posY + 95
            
            'Arquivo de Recuperação
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 5.5 scalefont setfont"
            'Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
            Print #outFN, posX & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
            Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
            'Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 5.5 scalefont setfont"
            Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (FIRMA/SEQ:.: " & Reg(i).SEQUENCIAL & ") show"

            posY = posY + 117
            
            'Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX + 1080 & " pontos -" & posY - 10 & " pontos moveto (" & Reg(i).Matrícula_Imovel & ") show"
            Print #outFN, posX + 1300 & " pontos -" & posY - 10 & " pontos moveto (" & Reg(i).MES_ANO & ") show"
                
            posY = posY + 180
            'Nome, Matricula
            Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
            Print #outFN, posX + 835 & " pontos -" & posY & " pontos moveto (" & Reg(i).Data_Vencimento & ") show"
            Print #outFN, posX + 1260 & " pontos -" & posY & " pontos moveto (" & Reg(i).Valor_Conta & ") show"
                
           'Linha Digitável e Código de Barras
'            posY = posY - 130
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 330 & " pontos -" & posY + 60 & " pontos moveto (" & Reg(i).Representacao_numerica_codigo_barras & ") show"
            
            posY = posY + 240
            Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
            Print #outFN, posX + 115 & " pontos -" & posY & " pontos moveto (" & Reg(i).Codigo_Barras & ") show"
                
        Next
        Print #outFN, "showpage"
        
        If (frmPrincipal.chkImprimirVerso.Value = 1) Then
            Print #outFN, "save EPSForm2 execform restore"
            Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
            Print #outFN, "showpage"
        End If
'        For i = 2 To 1 Step -1
'            If Len(Trim(Reg(i).NOME_CONTA)) > 0 Then totRegs = totRegs + 1
'            Select Case i
'                Case 1
'                  posX = 2070
'                  posY = 1880
'                Case 2
'                  posX = 320
'                  posY = 1890
'            End Select
'
'            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7.5 scalefont setfont"
'            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME_CONTA & ") show"
'            Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
'            posY = posY + 40
'            Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO_IMOVEL & Reg(i).Endereco_Imovel2 & ") show"
'
'            Print #outFN, posX & " pontos -" & posY + 130 & " pontos moveto (Grupo: " & Reg(i).Grupo_Faturamento & ") show"
'            Print #outFN, posX + 300 & " pontos -" & posY + 130 & " pontos moveto (Sequencial: " & Reg(i).SEQUENCIAL & ") show"
'
'            Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
'            Print #outFN, posX & " pontos -" & posY + 215 & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
'            Print #outFN, posX + 300 & " pontos -" & posY + 215 & " pontos moveto (Arq.: " & wArq & ") show"
'         Next
'
'         Print #outFN, "showpage"
               
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

       'If wRec = 10 Then End
End Sub
Private Function CalcularBarra(ByVal wMaior As Long, ByVal wConsumo As Long) As Long
        
        Dim wBigPoint As Integer 'Maior ponto que a barra pode alcançar
        wBigPoint = 50
        
        Dim wPoint As Long
        
        wPoint = CLng((CLng(wConsumo) * CLng(wBigPoint)) / wMaior)
        
        CalcularBarra = wPoint

End Function
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer, wCont As Integer, wIndice As Integer, wIndiceConsumo As Integer, wIndiceMensagem As Integer, wContRegs As Integer
        Dim wFlagDetalhe As Boolean
        Dim wQtdLinha As Double
        Dim wQtdAtual As Double
        wRec = 0
        wCont = 0
        wIndice = 1
        wQtdFaturas(1) = 0
        wQtdFaturas(2) = 0
        wIndiceConsumo = 0
        wIndiceMensagem = 0
        wContRegs = 1
        wFlagDetalhe = False
        totRegs = 0
        wQtdLinha = 0
        wQtdAtual = 0

        'Dividir Registros
        Open pastaProcessar & "temporario.tmp" For Output As #23
        Do While Not EOF(inFN)
           DoEvents
           ler inFN, linha
           Print #23, linha
           wQtdLinha = wQtdLinha + 1
        Loop
        
        Close #23
        
'        'Dividir Registros
'        Open pastaProcessar & "temporario1.tmp" For Output As #24
'        Open pastaProcessar & "temporario2.tmp" For Output As #25
'        Open pastaProcessar & "temporario.tmp" For Input As #23
'        Do While Not EOF(23)
'           DoEvents
'           Line Input #23, linha
'           wQtdAtual = wQtdAtual + 1
'
'           If wQtdAtual >= ((wQtdLinha / 2) + 1) Then
'              Print #25, linha
'           Else
'              Print #24, linha
'           End If
'        Loop
        
        Close #23
'        Close #24
'        Close #25
        Close #inFN

        totRegs = 0
        wCont = 0

        'Gravar Cabeçalho
        GravaHeader

        wArq = frmPrincipal.LstSpool.List(frmPrincipal.LstSpool.ListIndex)
        Open pastaProcessar & "temporario.tmp" For Input As #29

        Do While Not EOF(29)
           DoEvents
           For i = 1 To 2
'                If i = 1 Then
'                   Line Input #24, linha
'                Else
'                   Line Input #25, linha
'                End If
                
                If EOF(29) Then
                    Exit Do
                End If
                
                Line Input #29, linha
                                
                'linha = Replace(linha, "(", " ")
                'linha = Replace(linha, ")", " ")
                linha = Replace(linha, "[", " ")
                linha = Replace(linha, "]", " ")
                
                'linha = Replace(linha, "(", " ")
                'linha = Replace(linha, ")", " ")
                linha = Replace(linha, "[", " ")
                linha = Replace(linha, "]", " ")
                
                Reg(i).MENSAGEM = Mid(linha, 1, 50)
                Reg(i).Nome_Localidade = Mid(linha, 51, 30)
                Reg(i).Matrícula_Imovel = Mid(linha, 81, 10)
                Reg(i).MES_ANO = Mid(linha, 91, 9)
                Reg(i).NOME_CONTA = Mid(linha, 100, 50)
                Reg(i).ENDERECO_IMOVEL = Mid(linha, 150, 50)
                Reg(i).Endereco_Imovel2 = Mid(linha, 200, 50)
                Reg(i).INSCRICAO = Mid(linha, 250, 20)
                Reg(i).Grupo_Faturamento = Mid(linha, 270, 2)
                Reg(i).Codigo_debito_automatico = Mid(linha, 272, 10)
                Reg(i).Codigo_cliente_responsável = Mid(linha, 282, 8)
                Reg(i).Endereco_correspondencia_responsavel = Mid(linha, 290, 50)
                Reg(i).Endereco_correspondencia_responsavel2 = Mid(linha, 340, 50)
                Reg(i).situacao_ligacao_agua = Mid(linha, 390, 20)
                Reg(i).situacao_ligacao_esgoto = Mid(linha, 410, 20)
                Reg(i).Economias_Residencial = Mid(linha, 430, 4)
                Reg(i).Economias_Comercial = Mid(linha, 434, 4)
                Reg(i).Economias_Industrial = Mid(linha, 438, 4)
                Reg(i).Economias_Publico = Mid(linha, 442, 4)
                Reg(i).Hidrometro = Mid(linha, 446, 10)
                Reg(i).Data_Leitura_Anterior = Mid(linha, 456, 10)
                Reg(i).Data_Leitura_Atual = Mid(linha, 466, 10)
                Reg(i).Tipo_Consumo = Mid(linha, 476, 41)
                Reg(i).Leitura_Anterior_Agua = Mid(linha, 517, 7)
                Reg(i).Constante_Consumo_Faturado_Agua = Mid(linha, 524, 8)
                Reg(i).Consumo_Faturado_Agua = Mid(linha, 532, 5)
                Reg(i).Leitura_Atual_Agua = Mid(linha, 537, 7)
                Reg(i).Anormalidade_Leitura_Agua = Mid(linha, 544, 25)
                Reg(i).Leitura_Faturada_Agua = Mid(linha, 569, 7)
                Reg(i).Anormalidade_Consumo_Agua = Mid(linha, 576, 25)
                Reg(i).Leitura_Anterior_Esgoto = Mid(linha, 601, 7)
                Reg(i).Constante_Consumo_Faturado_Esgoto = Mid(linha, 608, 8)
                Reg(i).Consumo_Faturado_Esgoto = Mid(linha, 616, 5)
                Reg(i).Leitura_Atual_Esgoto = Mid(linha, 621, 7)
                Reg(i).Anormalidade_Leitura_Esgoto = Mid(linha, 628, 25)
                Reg(i).Leitura_Fat_Esgoto = Mid(linha, 653, 7)
                Reg(i).Anormalidade_Consumo_Esgoto = Mid(linha, 660, 25)
                Reg(i).Consumo_Mes1 = Mid(linha, 685, 19)
                Reg(i).Consumo_Mes2 = Mid(linha, 704, 19)
                Reg(i).Consumo_Mes3 = Mid(linha, 723, 19)
                Reg(i).Consumo_Mes4 = Mid(linha, 742, 19)
                Reg(i).Consumo_Mes5 = Mid(linha, 761, 19)
                Reg(i).Consumo_Mes6 = Mid(linha, 780, 19)
                Reg(i).A1 = Mid(linha, 693, 5)
                Reg(i).A2 = Mid(linha, 699, 5)
                Reg(i).B1 = Mid(linha, 712, 5)
                Reg(i).B2 = Mid(linha, 718, 5)
                Reg(i).C1 = Mid(linha, 731, 5)
                Reg(i).C2 = Mid(linha, 737, 5)
                Reg(i).D1 = Mid(linha, 750, 5)
                Reg(i).D2 = Mid(linha, 756, 5)
                Reg(i).E1 = Mid(linha, 769, 5)
                Reg(i).E2 = Mid(linha, 775, 5)
                Reg(i).F1 = Mid(linha, 788, 5)
                Reg(i).F2 = Mid(linha, 794, 5)
                Reg(i).M1 = Mid(linha, 806, 5)
                Reg(i).M2 = Mid(linha, 812, 5)
                Reg(i).MEDIA_CONSUMO = Mid(linha, 799, 19)
                Reg(i).Turbidez_Exigida = Mid(linha, 818, 5)
                Reg(i).Turbidez_Realizada = Mid(linha, 823, 5)
                Reg(i).Turbidez_Conforme = Mid(linha, 828, 5)
                Reg(i).Cor_Aparente_Exigida = Mid(linha, 833, 5)
                Reg(i).Cor_Aparente_Realizada = Mid(linha, 838, 5)
                Reg(i).Cor_Aparente_Conforme = Mid(linha, 843, 5)
                Reg(i).Cloro_Residual_Livre_Exigida = Mid(linha, 848, 5)
                Reg(i).Cloro_Residual_Livre_Realizada = Mid(linha, 853, 5)
                Reg(i).Cloro_Residual_Livre_Conforme = Mid(linha, 858, 5)
                Reg(i).Coliformes_Totais_Exigida = Mid(linha, 863, 5)
                Reg(i).Coliformes_Totais_Realizada = Mid(linha, 868, 5)
                Reg(i).Coliformes_Totais_Conforme = Mid(linha, 873, 5)
                Reg(i).E_Coli_Exigida = Mid(linha, 878, 5)
                Reg(i).E_Coli_Realizada = Mid(linha, 883, 5)
                Reg(i).E_Coli_Conforme = Mid(linha, 888, 5)
                Reg(i).Observacao1 = Mid(linha, 893, 62)
                Reg(i).Observacao2 = Mid(linha, 955, 120)
                Reg(i).Observacao3 = Mid(linha, 1075, 100)
                Reg(i).Servicos1 = Mid(linha, 1175, 50)
                Reg(i).Consumo_Faixa1 = Mid(linha, 1225, 15)
                Reg(i).Valor1 = Mid(linha, 1240, 14)
                Reg(i).Servicos2 = Mid(linha, 1254, 50)
                Reg(i).Consumo_Faixa2 = Mid(linha, 1304, 15)
                Reg(i).Valor2 = Mid(linha, 1319, 14)
                Reg(i).Servicos3 = Mid(linha, 1333, 50)
                Reg(i).Consumo_Faixa3 = Mid(linha, 1383, 15)
                Reg(i).Valor3 = Mid(linha, 1398, 14)
                Reg(i).Servicos4 = Mid(linha, 1412, 50)
                Reg(i).Consumo_Faixa4 = Mid(linha, 1462, 15)
                Reg(i).Valor4 = Mid(linha, 1477, 14)
                Reg(i).Servicos5 = Mid(linha, 1491, 50)
                Reg(i).Consumo_Faixa5 = Mid(linha, 1541, 15)
                Reg(i).Valor5 = Mid(linha, 1556, 14)
                Reg(i).Servicos6 = Mid(linha, 1570, 50)
                Reg(i).Consumo_Faixa6 = Mid(linha, 1620, 15)
                Reg(i).Valor6 = Mid(linha, 1635, 14)
                Reg(i).Servicos7 = Mid(linha, 1649, 50)
                Reg(i).Consumo_Faixa7 = Mid(linha, 1699, 15)
                Reg(i).Valor7 = Mid(linha, 1714, 14)
                Reg(i).Servicos8 = Mid(linha, 1728, 50)
                Reg(i).Consumo_Faixa8 = Mid(linha, 1778, 15)
                Reg(i).Valor8 = Mid(linha, 1793, 14)
                Reg(i).Servicos9 = Mid(linha, 1807, 50)
                Reg(i).Consumo_Faixa9 = Mid(linha, 1857, 15)
                Reg(i).Valor9 = Mid(linha, 1872, 14)
                Reg(i).Servicos10 = Mid(linha, 1886, 50)
                Reg(i).Consumo_Faixa10 = Mid(linha, 1936, 15)
                Reg(i).Valor10 = Mid(linha, 1951, 14)
                Reg(i).Servicos11 = Mid(linha, 1965, 50)
                Reg(i).Consumo11 = Mid(linha, 2015, 15)
                Reg(i).Valor11 = Mid(linha, 2030, 14)
                Reg(i).Servicos12 = Mid(linha, 2044, 50)
                Reg(i).Consumo12 = Mid(linha, 2094, 15)
                Reg(i).Valor12 = Mid(linha, 2109, 14)
                Reg(i).Servicos13 = Mid(linha, 2123, 50)
                Reg(i).Consumo_Faixa13 = Mid(linha, 2173, 15)
                Reg(i).Valor13 = Mid(linha, 2188, 14)
                Reg(i).Servicos14 = Mid(linha, 2202, 50)
                Reg(i).Consumo_Faixa14 = Mid(linha, 2252, 15)
                Reg(i).Valor14 = Mid(linha, 2267, 14)
                Reg(i).Servicos15 = Mid(linha, 2281, 50)
                Reg(i).ConsumoFaixa15 = Mid(linha, 2331, 15)
                Reg(i).Valor15 = Mid(linha, 2346, 14)
                Reg(i).Mensagem_rateio_consumo_fixo_esgoto = Mid(linha, 2360, 31)
                Reg(i).PIS_Base_Calculo = Mid(linha, 2391, 14)
                Reg(i).PIS_Percentual = Mid(linha, 2405, 5)
                Reg(i).PIS_Valor_Imposto = Mid(linha, 2410, 14)
                Reg(i).COFINS_Base_Calculo = Mid(linha, 2424, 14)
                Reg(i).COFINS_Percentual = Mid(linha, 2438, 5)
                Reg(i).COFINS_Valor_Imposto = Mid(linha, 2443, 14)
                Reg(i).Data_Vencimento = Mid(linha, 2457, 10)
                Reg(i).Valor_Conta = Mid(linha, 2467, 14)
                Reg(i).Mensagem_conta1 = Mid(linha, 2481, 100)
                Reg(i).Mensagem_conta2 = Mid(linha, 2581, 100)
                Reg(i).Mensagem_conta3 = Mid(linha, 2681, 100)
                Reg(i).SEQUENCIAL = Mid(linha, 3004, 10)
                Reg(i).Representacao_numerica_codigo_barras = Mid(linha, 2788, 55)
                Reg(i).Codigo_Barras = Mid(linha, 2843, 116)
                Reg(i).Matrícula_Imovel = Mid(linha, 2959, 10)
                Reg(i).VENCIMENTO = Mid(linha, 2978, 10)

'                Reg(i).Cedente = Mid(linha, 2908, 50)
'                Reg(i).Agencia_Cedente = Mid(linha, 2958, 13)
'                Reg(i).Data_Documento = Mid(linha, 2971, 10)
'                Reg(i).Numero_Documento = Mid(linha, 2981, 10)
'                Reg(i).Especie_Documento = Mid(linha, 2991, 3)
'                Reg(i).Aceite = Mid(linha, 2994, 1)
'                Reg(i).Data_Processamento = Mid(linha, 2995, 10)
'                Reg(i).Nosso_Numero = Mid(linha, 3005, 19)
'                Reg(i).Carteira = Mid(linha, 3024, 2)
'                Reg(i).Valor_Documento = Mid(linha, 3026, 14)
'                Reg(i).Sacado_linha_1a = Mid(linha, 3040, 50)
'                Reg(i).Sacado_linha1b = Mid(linha, 3090, 16)
'                Reg(i).Sacado_linha1c = Mid(linha, 3106, 10)
'                Reg(i).Sacado_linha1d = Mid(linha, 3116, 14)
'                Reg(i).Sacado_linha1e = Mid(linha, 3130, 9)
'                Reg(i).Sacado_linha2a = Mid(linha, 3139, 50)
'                Reg(i).Sacado_linha3a = Mid(linha, 3189, 50)

                 If Len(Trim(Reg(i).INSCRICAO)) Then totRegs = totRegs + 1
              
                frmPrincipal.txQtdReg.Text = totRegs

                
                If i = 2 Then
                   MontaExtrato
                   LimpaReg
                End If
           
            Next
        
        Loop
        If wIndice > 1 Then MontaExtrato
        Screen.MousePointer = vbDefault

        frmResultados.txtTotCli_STL.Text = totRegs
        frmResultados.txtTotReg_STL.Text = wRec
        frmResultados.txArquivo.Text = wArq

        frmResultados.txtAplicacao = AppDesc

        Close #outFN
        Close #29

    '    Kill pastaProcessar & "*.tmp"

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
        Open App.Path & "\Conta_Int.ps" For Input As #200
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
        '---------------------------------------------------------------
        'Segunda Imagem
        Print #outFN, "/ImageData2"
        Print #outFN, "currentfile"
        Print #outFN, "<< /Filter /SubFileDecode"
        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
        Print #outFN, ">> /ReusableStreamDecode filter"
        Open App.Path & "\Conta_Ext.ps" For Input As #200
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

        'Terceira Imagem Imagem
        Print #outFN, "/ImageData3"
        Print #outFN, "currentfile"
        Print #outFN, "<< /Filter /SubFileDecode"
        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
        Print #outFN, ">> /ReusableStreamDecode filter"
        Open App.Path & "\Conta_Int_Responsavel.ps" For Input As #200
        While Not EOF(200)
            Line Input #200, wCabeca
            Print #outFN, wCabeca
        Wend
        Close #200
        Print #outFN, "*EOD*"
        Print #outFN, "def"
        Print #outFN, " "
        Print #outFN, "/EPSForm3"
        Print #outFN, "<< /FormType 1"
        Print #outFN, "   /BBox [0 0 595 900]"
        Print #outFN, "   /Matrix [ 1 0 0 1 0 0]"
        Print #outFN, "   /PaintProc"
        Print #outFN, "   { pop"
        Print #outFN, "       /ostate save def"
        Print #outFN, "         /showpage {} def"
        Print #outFN, "         /setpagedevice /pop load def"
        Print #outFN, "         ImageData3 0 setfileposition ImageData3 cvx exec"
        Print #outFN, "       ostate restore"
        Print #outFN, "   } bind"
        Print #outFN, ">> def"
        Print #outFN, "%% Fim da Terceira Imagem"
        Print #outFN, " "

        Print #outFN, "%%============================"
        Print #outFN, "%% Inicio dos Dados Variaveis"
        Print #outFN, "%%============================"
        Print #outFN, " "
        Print #outFN, "%Aplicação: CONTA COMPESA A5"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        For i = 1 To 2
                Reg(i).MENSAGEM = ""
                Reg(i).Nome_Localidade = ""
                Reg(i).Matrícula_Imovel = ""
                Reg(i).MES_ANO = ""
                Reg(i).NOME_CONTA = ""
                Reg(i).ENDERECO_IMOVEL = ""
                Reg(i).Endereco_Imovel2 = ""
                Reg(i).INSCRICAO = ""
                Reg(i).Grupo_Faturamento = ""
                Reg(i).Codigo_debito_automatico = ""
                Reg(i).Codigo_cliente_responsável = ""
                Reg(i).Endereco_correspondencia_responsavel = ""
                Reg(i).Endereco_correspondencia_responsavel2 = ""
                Reg(i).situacao_ligacao_agua = ""
                Reg(i).situacao_ligacao_esgoto = ""
                Reg(i).Economias_Residencial = ""
                Reg(i).Economias_Comercial = ""
                Reg(i).Economias_Industrial = ""
                Reg(i).Economias_Publico = ""
                Reg(i).Hidrometro = ""
                Reg(i).Data_Leitura_Anterior = ""
                Reg(i).Data_Leitura_Atual = ""
                Reg(i).Tipo_Consumo = ""
                Reg(i).Leitura_Anterior_Agua = ""
                Reg(i).Constante_Consumo_Faturado_Agua = ""
                Reg(i).Consumo_Faturado_Agua = ""
                Reg(i).Leitura_Atual_Agua = ""
                Reg(i).Anormalidade_Leitura_Agua = ""
                Reg(i).Leitura_Faturada_Agua = ""
                Reg(i).Anormalidade_Consumo_Agua = ""
                Reg(i).Leitura_Anterior_Esgoto = ""
                Reg(i).Constante_Consumo_Faturado_Esgoto = ""
                Reg(i).Consumo_Faturado_Esgoto = ""
                Reg(i).Leitura_Atual_Esgoto = ""
                Reg(i).Anormalidade_Leitura_Esgoto = ""
                Reg(i).Leitura_Fat_Esgoto = ""
                Reg(i).Anormalidade_Consumo_Esgoto = ""
                Reg(i).Consumo_Mes1 = ""
                Reg(i).Consumo_Mes2 = ""
                Reg(i).Consumo_Mes3 = ""
                Reg(i).Consumo_Mes4 = ""
                Reg(i).Consumo_Mes5 = ""
                Reg(i).Consumo_Mes6 = ""
                
                
                Reg(i).A1 = ""
                Reg(i).A2 = ""
                Reg(i).B1 = ""
                Reg(i).B2 = ""
                Reg(i).C1 = ""
                Reg(i).C2 = ""
                Reg(i).D1 = ""
                Reg(i).D2 = ""
                Reg(i).E1 = ""
                Reg(i).E2 = ""
                Reg(i).F1 = ""
                Reg(i).F2 = ""
                Reg(i).M1 = ""
                Reg(i).M2 = ""
                          
                         
                Reg(i).MEDIA_CONSUMO = ""
                Reg(i).Turbidez_Exigida = ""
                Reg(i).Turbidez_Realizada = ""
                Reg(i).Turbidez_Conforme = ""
                Reg(i).Cor_Aparente_Exigida = ""
                Reg(i).Cor_Aparente_Realizada = ""
                Reg(i).Cor_Aparente_Conforme = ""
                Reg(i).Cloro_Residual_Livre_Exigida = ""
                Reg(i).Cloro_Residual_Livre_Realizada = ""
                Reg(i).Cloro_Residual_Livre_Conforme = ""
                Reg(i).Coliformes_Totais_Exigida = ""
                Reg(i).Coliformes_Totais_Realizada = ""
                Reg(i).Coliformes_Totais_Conforme = ""
                Reg(i).E_Coli_Exigida = ""
                Reg(i).E_Coli_Realizada = ""
                Reg(i).E_Coli_Conforme = ""
                Reg(i).Observacao1 = ""
                Reg(i).Observacao2 = ""
                Reg(i).Observacao3 = ""
                Reg(i).Servicos1 = ""
                Reg(i).Consumo_Faixa1 = ""
                Reg(i).Valor1 = ""
                Reg(i).Servicos2 = ""
                Reg(i).Consumo_Faixa2 = ""
                Reg(i).Valor2 = ""
                Reg(i).Servicos3 = ""
                Reg(i).Consumo_Faixa3 = ""
                Reg(i).Valor3 = ""
                Reg(i).Servicos4 = ""
                Reg(i).Consumo_Faixa4 = ""
                Reg(i).Valor4 = ""
                Reg(i).Servicos5 = ""
                Reg(i).Consumo_Faixa5 = ""
                Reg(i).Valor5 = ""
                Reg(i).Servicos6 = ""
                Reg(i).Consumo_Faixa6 = ""
                Reg(i).Valor6 = ""
                Reg(i).Servicos7 = ""
                Reg(i).Consumo_Faixa7 = ""
                Reg(i).Valor7 = ""
                Reg(i).Servicos8 = ""
                Reg(i).Consumo_Faixa8 = ""
                Reg(i).Valor8 = ""
                Reg(i).Servicos9 = ""
                Reg(i).Consumo_Faixa9 = ""
                Reg(i).Valor9 = ""
                Reg(i).Servicos10 = ""
                Reg(i).Consumo_Faixa10 = ""
                Reg(i).Valor10 = ""
                Reg(i).Servicos11 = ""
                Reg(i).Consumo11 = ""
                Reg(i).Valor11 = ""
                Reg(i).Servicos12 = ""
                Reg(i).Consumo12 = ""
                Reg(i).Valor12 = ""
                Reg(i).Servicos13 = ""
                Reg(i).Consumo_Faixa13 = ""
                Reg(i).Valor13 = ""
                Reg(i).Servicos14 = ""
                Reg(i).Consumo_Faixa14 = ""
                Reg(i).Valor14 = ""
                Reg(i).Servicos15 = ""
                Reg(i).ConsumoFaixa15 = ""
                Reg(i).Valor15 = ""
                Reg(i).Mensagem_rateio_consumo_fixo_esgoto = ""
                Reg(i).PIS_Base_Calculo = ""
                Reg(i).PIS_Percentual = ""
                Reg(i).PIS_Valor_Imposto = ""
                Reg(i).COFINS_Base_Calculo = ""
                Reg(i).COFINS_Percentual = ""
                Reg(i).COFINS_Valor_Imposto = ""
                Reg(i).Data_Vencimento = ""
                Reg(i).Valor_Conta = ""
                
                Reg(i).Mensagem_conta1 = ""
                Reg(i).Mensagem_conta2 = ""
                Reg(i).Mensagem_conta3 = ""
                Reg(i).SEQUENCIAL = ""
                Reg(i).Código_Banco = ""
                Reg(i).Representacao_numerica_codigo_barras = ""
                Reg(i).Local_Pagamento = ""
                Reg(i).VENCIMENTO = ""
                Reg(i).Cedente = ""
                Reg(i).Agencia_Cedente = ""
                Reg(i).Data_Documento = ""
                Reg(i).Numero_Documento = ""
                Reg(i).Especie_Documento = ""
                Reg(i).Aceite = ""
                Reg(i).Data_Processamento = ""
                Reg(i).Nosso_Numero = ""
                Reg(i).Carteira = ""
                Reg(i).Valor_Documento = ""
                Reg(i).Sacado_linha_1a = ""
                Reg(i).Sacado_linha1b = ""
                Reg(i).Sacado_linha1c = ""
                Reg(i).Sacado_linha1d = ""
                Reg(i).Sacado_linha1e = ""
                Reg(i).Sacado_linha2a = ""
                Reg(i).Sacado_linha3a = ""
                Reg(i).Codigo_Barras = ""
        Next
End Sub
