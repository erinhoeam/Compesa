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
'Desenvolvido por Sistemap Consultoria
'                 Analista de Sistemas
'Data: 27 de Novembro de 2008
'Contatos - Email: aziel@sistemap.com.br
'                  erinhoeam@hotmail.com
'           Fone: (81) 9832.6084
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
        Data_Vencimento   As String  ' 2892    2462    10
        Valor_Conta   As String  ' 2463    2476    14
        Mensagem_conta1   As String  '  2477    2576    100
        Mensagem_conta2   As String  '  2577    2676    100
        Mensagem_conta3   As String  ' 2677    2776    100
        SEQUENCIAL   As String  ' 2777    2783    7
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
        CPF_CNPJ As String '3365 -> 3383 - 18
        MSG01 As String '3370 a 3411
        MSG02 As String '3412 a 3341
        MSG03 As String '3342 a 3484
        MSG04 As String '3485 a 3545

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
        'Print #outFN, "save EPSForm1 execform restore"
       'Capturar o Maior Consumo
        Dim wMaior As Long
        Dim wDivisor As Integer
        wMaior = 0
        
        If Val(Reg.A1) > wMaior Then wMaior = Val(Reg.A1)
        If Val(Reg.A2) > wMaior Then wMaior = Val(Reg.A2)
        If Val(Reg.B1) > wMaior Then wMaior = Val(Reg.B1)
        If Val(Reg.B2) > wMaior Then wMaior = Val(Reg.B2)
        If Val(Reg.C1) > wMaior Then wMaior = Val(Reg.C1)
        If Val(Reg.C2) > wMaior Then wMaior = Val(Reg.C2)
        If Val(Reg.D1) > wMaior Then wMaior = Val(Reg.D1)
        If Val(Reg.D2) > wMaior Then wMaior = Val(Reg.D2)
        If Val(Reg.E1) > wMaior Then wMaior = Val(Reg.E1)
        If Val(Reg.E2) > wMaior Then wMaior = Val(Reg.E2)
        If Val(Reg.F1) > wMaior Then wMaior = Val(Reg.F1)
        If Val(Reg.F2) > wMaior Then wMaior = Val(Reg.F2)
        
        'Testes
         wDivisor = 1
         If wMaior >= 100 And wMaior < 200 Then wDivisor = 3
         If wMaior >= 200 And wMaior < 300 Then wDivisor = 3.5
         If wMaior >= 300 And wMaior < 400 Then wDivisor = 4.7
         If wMaior >= 400 And wMaior < 500 Then wDivisor = 6
         If wMaior >= 500 And wMaior < 600 Then wDivisor = 7
         If wMaior >= 600 And wMaior < 700 Then wDivisor = 8
         If wMaior >= 700 And wMaior < 800 Then wDivisor = 9
         If wMaior >= 800 And wMaior < 900 Then wDivisor = 11
         If wMaior >= 900 And wMaior < 1000 Then wDivisor = 11.5
         If wMaior >= 1000 And wMaior < 1100 Then wDivisor = 12
         If wMaior >= 1100 And wMaior < 1200 Then wDivisor = 12.5
         If wMaior >= 1200 And wMaior < 1300 Then wDivisor = 13.5
         If wMaior >= 1300 And wMaior < 1400 Then wDivisor = 14.5
         If wMaior >= 1400 And wMaior < 1500 Then wDivisor = 15.5
         If wMaior >= 1500 And wMaior < 1600 Then wDivisor = 18.7
         If wMaior >= 1600 And wMaior < 1700 Then wDivisor = 19
         If wMaior >= 1700 And wMaior < 1800 Then wDivisor = 19.5
         If wMaior >= 1800 And wMaior < 1900 Then wDivisor = 20
         If wMaior >= 1900 And wMaior < 2000 Then wDivisor = 20.5
         If wMaior >= 2000 And wMaior < 2100 Then wDivisor = 21.5
         If wMaior >= 2100 And wMaior < 2200 Then wDivisor = 22.5
         If wMaior >= 2200 And wMaior < 2300 Then wDivisor = 23.5
         If wMaior >= 2300 And wMaior < 2400 Then wDivisor = 24.5
         If wMaior >= 2400 And wMaior < 2500 Then wDivisor = 25.5
         If wMaior >= 2500 And wMaior < 2600 Then wDivisor = 26.5
         If wMaior >= 2600 And wMaior < 2700 Then wDivisor = 27.5
         If wMaior >= 2700 And wMaior < 2800 Then wDivisor = 28.5
         If wMaior >= 2800 And wMaior < 2900 Then wDivisor = 29.5
         If wMaior >= 2900 And wMaior < 3000 Then wDivisor = 30.5
         If wMaior >= 3000 And wMaior < 3100 Then wDivisor = 31.5
         If wMaior >= 3100 And wMaior < 3200 Then wDivisor = 32.5
         If wMaior >= 3200 And wMaior < 3300 Then wDivisor = 33.5
         If wMaior >= 3300 And wMaior < 3400 Then wDivisor = 34.5
         If wMaior >= 3500 And wMaior < 3500 Then wDivisor = 35.5
       
       '-------------------------------------------------------
        If Len(Trim(Reg.A1)) > 0 And wMaior > 0 Then
           Print #outFN, "%Primeira Barra"
           Print #outFN, "newpath"
           Print #outFN, "120 610 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.A1)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.A1 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
           Print #outFN, "-3 0 rlineto"
           Print #outFN, "closepath"
           Print #outFN, "0 0 0 setrgbcolor"
           Print #outFN, "1 setlinewidth"
           Print #outFN, "stroke"
        End If
        If Len(Trim(Reg.A2)) > 0 And wMaior > 0 Then
           Print #outFN, "120 607 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.A2)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.A2 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
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
        If Len(Trim(Reg.B1)) > 0 And wMaior > 0 Then
           Print #outFN, "%Segunda Barra"
           Print #outFN, "newpath"
           Print #outFN, "120 600 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.B1)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.B1 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
           Print #outFN, "-3 0 rlineto"
           Print #outFN, "closepath"
           Print #outFN, "0 0 0 setrgbcolor"
           Print #outFN, "1 setlinewidth"
           Print #outFN, "stroke"
        End If
        If Len(Trim(Reg.B2)) > 0 And wMaior > 0 Then
           Print #outFN, "120 597 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.B2)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.B2 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
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
        If Len(Trim(Reg.C1)) > 0 And wMaior > 0 Then
           Print #outFN, "%Terceira Barra"
           Print #outFN, "newpath"
           Print #outFN, "120 588 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.C1)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.C1 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
           Print #outFN, "-3 0 rlineto"
           Print #outFN, "closepath"
           Print #outFN, "0 0 0 setrgbcolor"
           Print #outFN, "1 setlinewidth"
           Print #outFN, "stroke"
        End If
        If Len(Trim(Reg.C2)) > 0 And wMaior > 0 Then
           Print #outFN, "120 585 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.C2)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.C2 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
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
        If Len(Trim(Reg.D1)) > 0 And wMaior > 0 Then
           Print #outFN, "%Quarta Barra"
           Print #outFN, "newpath"
           Print #outFN, "120 575 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.D1)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.D1 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
           Print #outFN, "-3 0 rlineto"
           Print #outFN, "closepath"
           Print #outFN, "0 0 0 setrgbcolor"
           Print #outFN, "1 setlinewidth"
           Print #outFN, "stroke"
        End If
        If Len(Trim(Reg.D2)) > 0 And wMaior > 0 Then
           Print #outFN, "120 572 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.D2)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.D2 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
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
        If Val(Trim(Reg.E1)) > 0 And wMaior > 0 Then
           Print #outFN, "%Quinta Barra"
           Print #outFN, "newpath"
           Print #outFN, "120 564 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.E1)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.E1 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
           Print #outFN, "-3 0 rlineto"
           Print #outFN, "closepath"
           Print #outFN, "0 0 0 setrgbcolor"
           Print #outFN, "1 setlinewidth"
           Print #outFN, "stroke"
        End If
        If Val(Trim(Reg.E2)) > 0 And wMaior > 0 Then
           Print #outFN, "120 561 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.E2)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.E2 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
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
        If Val(Trim(Reg.F1)) > 0 And wMaior > 0 Then
           Print #outFN, "%Sexta Barra"
           Print #outFN, "newpath"
           Print #outFN, "120 552 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.F1)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.F1 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
           Print #outFN, "-3 0 rlineto"
           Print #outFN, "closepath"
           Print #outFN, "0 0 0 setrgbcolor"
           Print #outFN, "1 setlinewidth"
           Print #outFN, "stroke"
        End If
        If Val(Trim(Reg.F2)) > 0 And wMaior > 0 Then
           Print #outFN, "120 549 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.F2)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.F2 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
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
        If Val(Trim(Reg.M1)) > 0 And wMaior > 0 Then
           Print #outFN, "%Sétima Barra"
           Print #outFN, "newpath"
           Print #outFN, "120 539 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.M1)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.M1 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
           Print #outFN, "-3 0 rlineto"
           Print #outFN, "closepath"
           Print #outFN, "0 0 0 setrgbcolor"
           Print #outFN, "1 setlinewidth"
           Print #outFN, "stroke"
        End If
        If Val(Trim(Reg.M2)) > 0 And wMaior > 0 Then
           Print #outFN, "120 536 moveto"
           Print #outFN, "0 3 rlineto"
           Print #outFN, CalcularBarra(wMaior, CLng(Reg.M2)) & " 0 rlineto"
           'Print #outFN, (Int(Reg.M2 / wDivisor)) & " 0 rlineto"
           Print #outFN, "0 -3 rlineto"
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

        Print #outFN, "0 0 0 setrgbcolor"
        If frmPrincipal.chkBoletoResp.Value = 0 Then
            Print #outFN, "save EPSForm1 execform restore"
        Else
            Print #outFN, "save EPSForm3 execform restore"
        End If
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        'Rota
        Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
        Print #outFN, posX + 220 & " pontos -" & posY - 90 & " pontos moveto (" & Reg.Nosso_Numero & ") show"
        Print #outFN, posX + 820 & " pontos -" & posY - 90 & " pontos moveto (" & Reg.Nome_Localidade & ") show"
        posY = posY + 30
        'Nome, Matricula
        Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
        Print #outFN, posX + 1580 & " pontos -" & posY & " pontos moveto (" & Reg.Matrícula_Imovel & ") show"
        Print #outFN, posX + 1820 & " pontos -" & posY & " pontos moveto (" & Reg.MES_ANO & ") show"
        posY = posY + 15
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 6.5 scalefont setfont"
        Print #outFN, posX - 50 & " pontos -" & posY & " pontos moveto (" & Reg.NOME_CONTA & ") show"
        posY = posY + 30
        Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
        Print #outFN, posX - 50 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO_IMOVEL & Reg.Endereco_Imovel2 & ") show"

        Print #outFN, posX + 180 & " pontos -" & posY + 38 & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
        Print #outFN, posX + 430 & " pontos -" & posY + 38 & " pontos moveto (CNPJ/CPF:) show"
        Print #outFN, "/Helvetica-Condensed-AP findfont 6.5 scalefont setfont"
        Print #outFN, posX + 570 & " pontos -" & posY + 38 & " pontos moveto (" & Reg.CPF_CNPJ & ") show"
        Print #outFN, posX + 1080 & " pontos -" & posY + 38 & " pontos moveto (" & Reg.Grupo_Faturamento & ") show"
        Print #outFN, posX + 1865 & " pontos -" & posY + 38 & " pontos moveto (" & Reg.Codigo_debito_automatico & ") show"

        posY = posY + 110
        Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
        Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg.Codigo_cliente_responsável & ") show"
        Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Trim(Reg.Endereco_correspondencia_responsavel) & " " & Trim(Reg.Endereco_correspondencia_responsavel2) & ") show"
        posY = posY + 90
        Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg.situacao_ligacao_agua & ") show"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.situacao_ligacao_esgoto & ") show"
        
        Print #outFN, posX + 900 & " pontos -" & posY + 8 & " pontos moveto (" & Reg.Economias_Residencial & ") show"
        Print #outFN, posX + 1200 & " pontos -" & posY + 8 & " pontos moveto (" & Reg.Economias_Comercial & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY + 8 & " pontos moveto (" & Reg.Economias_Industrial & ") show"
        Print #outFN, posX + 1850 & " pontos -" & posY + 8 & " pontos moveto (" & Reg.Economias_Publico & ") show"
        
        posY = posY + 75
        Print #outFN, posX + 30 & " pontos -" & posY & " pontos moveto (" & Reg.Hidrometro & ") show"
        Print #outFN, posX + 430 & " pontos -" & posY & " pontos moveto (" & Reg.Data_Leitura_Anterior & ") show"
        Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Reg.Data_Leitura_Atual & ") show"
        Print #outFN, posX + 1430 & " pontos -" & posY & " pontos moveto (" & Reg.Tipo_Consumo & ") show"
        
       'Leituras
        Print #outFN, "/Helvetica-Condensed-AP findfont 7.6 scalefont setfont"
        posY = posY + 65
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.Leitura_Anterior_Agua & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Reg.Leitura_Anterior_Esgoto & ") show"
        
        Print #outFN, posX + 680 & " pontos -" & posY & " pontos moveto (" & Reg.Consumo_Faturado_Agua & ") show"
        Print #outFN, posX + 1850 & " pontos -" & posY & " pontos moveto (" & Reg.Consumo_Faturado_Esgoto & ") show"
        
        posY = posY + 33
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.Leitura_Atual_Agua & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Reg.Leitura_Atual_Esgoto & ") show"
        
        posY = posY + 33
        Print #outFN, posX + 150 & " pontos -" & posY & " pontos moveto (" & Reg.Leitura_Faturada_Agua & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Reg.Leitura_Fat_Esgoto & ") show"
        
        Print #outFN, "/Helvetica-Condensed-AP findfont 5.4 scalefont setfont"
        Print #outFN, posX + 430 & " pontos -" & posY - 4 & " pontos moveto (" & Reg.Anormalidade_Consumo_Agua & ") show"
        
        'Consumo
        Print #outFN, "/Helvetica-Condensed-AP findfont 6 scalefont setfont"
        posY = posY + 110
        Print #outFN, posX + 20 & " pontos -" & posY & " pontos moveto (" & Reg.Consumo_Mes1 & ") show"
        Print #outFN, posX + 20 & " pontos -" & posY + 47 & " pontos moveto (" & Reg.Consumo_Mes2 & ") show"
        Print #outFN, posX + 20 & " pontos -" & posY + 97 & " pontos moveto (" & Reg.Consumo_Mes3 & ") show"
        Print #outFN, posX + 20 & " pontos -" & posY + 148 & " pontos moveto (" & Reg.Consumo_Mes4 & ") show"
        Print #outFN, posX + 20 & " pontos -" & posY + 202 & " pontos moveto (" & Reg.Consumo_Mes5 & ") show"
        Print #outFN, posX + 20 & " pontos -" & posY + 248 & " pontos moveto (" & Reg.Consumo_Mes6 & ") show"
        Print #outFN, posX + 20 & " pontos -" & posY + 300 & " pontos moveto (" & Reg.MEDIA_CONSUMO & ") show"
        
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 4 scalefont setfont"
        Print #outFN, posX + 260 & " pontos -" & posY - 8 & " pontos moveto (A) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 7 & " pontos moveto (E) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 35 & " pontos moveto (A) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 50 & " pontos moveto (E) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 84 & " pontos moveto (A) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 99 & " pontos moveto (E) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 136 & " pontos moveto (A) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 150 & " pontos moveto (E) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 181 & " pontos moveto (A) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 197 & " pontos moveto (E) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 234 & " pontos moveto (A) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 249 & " pontos moveto (E) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 288 & " pontos moveto (A) show"
        Print #outFN, posX + 260 & " pontos -" & posY + 303 & " pontos moveto (E) show"
        
        
        'Turbidez
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        posY = posY + 46
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.Turbidez_Exigida & ") show"
        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (" & Reg.Turbidez_Realizada & ") show"
        Print #outFN, posX + 1900 & " pontos -" & posY & " pontos moveto (" & Reg.Turbidez_Conforme & ") show"
        
        'Cor Aparente
        posY = posY + 30
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.Cor_Aparente_Exigida & ") show"
        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (" & Reg.Cor_Aparente_Realizada & ") show"
        Print #outFN, posX + 1900 & " pontos -" & posY & " pontos moveto (" & Reg.Cor_Aparente_Conforme & ") show"
        
        'Cloro Residual
        posY = posY + 30
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.Cloro_Residual_Livre_Exigida & ") show"
        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (" & Reg.Cloro_Residual_Livre_Realizada & ") show"
        Print #outFN, posX + 1900 & " pontos -" & posY & " pontos moveto (" & Reg.Cloro_Residual_Livre_Conforme & ") show"
        
        'Cloro Residual
        posY = posY + 30
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.Coliformes_Totais_Exigida & ") show"
        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (" & Reg.Coliformes_Totais_Realizada & ") show"
        Print #outFN, posX + 1900 & " pontos -" & posY & " pontos moveto (" & Reg.Coliformes_Totais_Conforme & ") show"
        
        'E Coli
        posY = posY + 30
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.E_Coli_Exigida & ") show"
        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (" & Reg.E_Coli_Realizada & ") show"
        Print #outFN, posX + 1900 & " pontos -" & posY & " pontos moveto (" & Reg.E_Coli_Conforme & ") show"
                
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 5.6 scalefont setfont"
        Print #outFN, posX + 850 & " pontos -" & posY + 45 & " pontos moveto (" & Reg.Observacao1 & ") show"
        Print #outFN, posX + 628 & " pontos -" & posY + 75 & " pontos moveto (" & Reg.Observacao2 & ") show"
        Print #outFN, posX + 628 & " pontos -" & posY + 105 & " pontos moveto (" & Reg.Observacao3 & ") show"
                
       'DESCRIÇÃO DOS SERVIÇOS
        posY = posY + 255
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 7 scalefont setfont"
        
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos1 & ") show"
        Print #outFN, posX + 1320 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa1)) & Reg.Consumo_Faixa1 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor1)) & Reg.Valor1 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos2 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa2)) & Reg.Consumo_Faixa2 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor2)) & Reg.Valor2 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos3 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa3)) & Reg.Consumo_Faixa3 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor3)) & Reg.Valor3 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos4 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa4)) & Reg.Consumo_Faixa4 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor4)) & Reg.Valor4 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos5 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa5)) & Reg.Consumo_Faixa5 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor5)) & Reg.Valor5 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos6 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa6)) & Reg.Consumo_Faixa6 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor6)) & Reg.Valor6 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos7 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa7)) & Reg.Consumo_Faixa7 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor7)) & Reg.Valor7 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos8 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa8)) & Reg.Consumo_Faixa8 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor8)) & Reg.Valor8 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos9 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa9)) & Reg.Consumo_Faixa9 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor9)) & Reg.Valor9 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos10 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa10)) & Reg.Consumo_Faixa10 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor10)) & Reg.Valor10 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos11 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo11)) & Reg.Consumo11 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor11)) & Reg.Valor11 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos12 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo12)) & Reg.Consumo12 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor12)) & Reg.Valor12 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos13 & ") show"
        Print #outFN, posX + 1320 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa13)) & Reg.Consumo_Faixa13 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor13)) & Reg.Valor13 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos14 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.Consumo_Faixa14)) & Reg.Consumo_Faixa14 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor14)) & Reg.Valor14 & ") show"
        posY = posY + 30
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Servicos15 & ") show"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.ConsumoFaixa15)) & Reg.ConsumoFaixa15 & ") show"
        Print #outFN, posX + 1530 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.Valor15)) & Reg.Valor15 & ") show"
        
        posY = posY + 50
        Print #outFN, posX - 20 & " pontos -" & posY & " pontos moveto (" & Reg.Mensagem_rateio_consumo_fixo_esgoto & ") show"
        
        posY = posY + 70
        Print #outFN, "/LetterGothic12PitchBT-Roman-AP findfont 6.5 scalefont setfont"
        Print #outFN, posX + 520 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.PIS_Base_Calculo)) & Reg.PIS_Base_Calculo & ") show"
        Print #outFN, posX + 1020 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.PIS_Percentual)) & Reg.PIS_Percentual & ") show"
        Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.PIS_Valor_Imposto)) & Reg.PIS_Valor_Imposto & ") show"
        
        posY = posY + 30
        Print #outFN, posX + 520 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.COFINS_Base_Calculo)) & Reg.COFINS_Base_Calculo & ") show"
        Print #outFN, posX + 1020 & " pontos -" & posY & " pontos moveto (" & Space(15 - Len(Reg.COFINS_Percentual)) & Reg.COFINS_Percentual & ") show"
        Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Space(28 - Len(Reg.COFINS_Valor_Imposto)) & Reg.COFINS_Valor_Imposto & ") show"
        
        posY = posY + 70
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
        Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Reg.Data_Vencimento & ") show"
        Print #outFN, posX + 1790 & " pontos -" & posY & " pontos moveto (" & Reg.Valor_Conta & ") show"
        
        posY = posY + 92
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.Mensagem_conta1 & ") show"
        Print #outFN, posX & " pontos -" & posY + 40 & " pontos moveto (" & Reg.Mensagem_conta2 & ") show"
        Print #outFN, posX & " pontos -" & posY + 80 & " pontos moveto (" & Reg.Mensagem_conta3 & ") show"
        
        posY = posY + 115
            
        'Arquivo de Recuperação
        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 5.5 scalefont setfont"
        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (SEQ.: " & Reg.SEQUENCIAL & ") show"
        
        posY = posY + 70
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 12 scalefont setfont"
        Print #outFN, posX + 630 & " pontos -" & posY & " pontos moveto (" & Reg.Representacao_numerica_codigo_barras & ") show"
        posY = posY + 93
        Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
'        Print #outFN, posX - 10 & " pontos -" & posY & " pontos moveto (" & Reg.Local_Pagamento & ") show"
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Reg.VENCIMENTO & ") show"
        Print #outFN, "/Helvetica-Condensed-AP findfont 7.5 scalefont setfont"
        posY = posY + 50
        Print #outFN, posX - 10 & " pontos -" & posY & " pontos moveto (" & Reg.Cedente & ") show"
'        Print #outFN, posX + 1460 & " pontos -" & posY + 15 & " pontos moveto (" & Reg.Agencia_Cedente & ") show"
        posY = posY + 70
        Print #outFN, posX - 10 & " pontos -" & posY & " pontos moveto (" & Reg.Data_Documento & ") show"
        Print #outFN, posX + 340 & " pontos -" & posY & " pontos moveto (" & Reg.Numero_Documento & ") show"
        Print #outFN, posX + 720 & " pontos -" & posY & " pontos moveto (" & Reg.Especie_Documento & ") show"
'        Print #outFN, posX + 950 & " pontos -" & posY & " pontos moveto (" & Reg.Aceite & ") show"
        Print #outFN, posX + 1195 & " pontos -" & posY & " pontos moveto (" & Reg.Data_Processamento & ") show"
        Print #outFN, posX + 1660 & " pontos -" & posY + 45 & " pontos moveto (" & Reg.Nosso_Numero & ") show"
'        Print #outFN, posX - 10 & " pontos -" & posY - 400 & " pontos moveto (" & Reg.Nosso_Numero & ") show"
        posY = posY + 75
'       Print #outFN, posX + 340 & " pontos -" & posY & " pontos moveto (" & Reg.Carteira & ") show"
        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & "R$" & ") show"
        posY = posY + 70
        Print #outFN, "/LetterGothic12PitchBT-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX + 1280 & " pontos -" & posY & " pontos moveto (" & Space(30 - Len(Reg.Valor_Documento)) & Reg.Valor_Documento & ") show"
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        posY = posY - 10
        Print #outFN, posX - 70 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.MSG01 & ") show"
        posY = posY + 45
        Print #outFN, posX - 70 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.MSG02 & ") show"
        posY = posY + 45
        Print #outFN, posX - 70 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.MSG03 & ") show"
        posY = posY + 45
        Print #outFN, posX - 70 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.MSG04 & ") show"
        posY = posY + 210
        Print #outFN, posX - 50 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.Sacado_linha_1a & ") show"
        'Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        posY = posY + 45
        Print #outFN, posX - 50 & " pontos -" & posY + 30 & " pontos moveto (" & Reg.ENDERECO_IMOVEL & Reg.Endereco_Imovel2 & ") show"
        
        Print #outFN, posX + 930 & " pontos -" & posY - 71 & " pontos moveto (" & Reg.Sacado_linha1c & ") show"
        Print #outFN, posX + 1330 & " pontos -" & posY - 71 & " pontos moveto (" & Reg.Sacado_linha1e & ") show"
        posY = posY + 345
        Print #outFN, "/I2501P findfont 37.5 scalefont setfont"
        Print #outFN, posX - 50 & " pontos -" & posY & " pontos moveto (" & Reg.Codigo_Barras & ") show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        'Fim da Parte Interna
        Print #outFN, " "
        
        If (frmPrincipal.chkImprimirVerso.Value = 1) Then
        
        Print #outFN, "save EPSForm2 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
'
'        posX = 150
'        posY = 1850
'
'       'Endereçamento
'        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 9 scalefont setfont"
'        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (" & Reg.NOME_CONTA & ") show"
'        posY = posY + 80
'        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
'        If Len(Trim(Reg.Endereco_correspondencia_responsavel)) > 0 Then
'           Print #outFN, posX + 400 & " pontos -" & posY - 40 & " pontos moveto (" & Reg.Endereco_correspondencia_responsavel & Reg.Endereco_correspondencia_responsavel2 & ") show"
'        Else
'           Print #outFN, posX + 400 & " pontos -" & posY - 40 & " pontos moveto (" & Reg.ENDERECO_IMOVEL & Reg.Endereco_Imovel2 & ") show"
'        End If
'
'        posY = posY + 150
'        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
'        Print #outFN, posX + 400 & " pontos -" & posY & " pontos moveto (Grupo: " & Reg.Grupo_Faturamento & "     Sequencial: " & Reg.SEQUENCIAL & ") show"
'
'        posY = posY + 150
'        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
'        Print #outFN, posX + 1550 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
'        Print #outFN, posX + 1750 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
'
        Print #outFN, "showpage"
        End If
        'Fim da Parte Interna
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "


       'End
End Sub
Private Function CalcularBarra(ByVal wMaior As Long, ByVal wConsumo As Long) As Long
        
        Dim wBigPoint As Integer 'Maior ponto que a barra pode alcançar
        wBigPoint = 79
        
        Dim wPoint As Long
        
        wPoint = CLng((CLng(wConsumo) * CLng(wBigPoint)) / wMaior)
        
        CalcularBarra = wPoint

End Function

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
           
           'linha = Replace(linha, "(", " ")
           'linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")
           
           Reg.MENSAGEM = Mid(linha, 1, 50)
           Reg.Nome_Localidade = Mid(linha, 51, 30)
           Reg.Matrícula_Imovel = Mid(linha, 81, 10)
           Reg.MES_ANO = Mid(linha, 91, 9)
           Reg.NOME_CONTA = Mid(linha, 100, 50)
           Reg.ENDERECO_IMOVEL = Mid(linha, 150, 50)
           Reg.Endereco_Imovel2 = Mid(linha, 200, 50)
           Reg.INSCRICAO = Mid(linha, 250, 20)
           Reg.Grupo_Faturamento = Mid(linha, 270, 2)
           Reg.Codigo_debito_automatico = Mid(linha, 272, 10)
           Reg.Codigo_cliente_responsável = Mid(linha, 282, 8)
           Reg.Endereco_correspondencia_responsavel = Mid(linha, 290, 50)
           Reg.Endereco_correspondencia_responsavel2 = Mid(linha, 340, 50)
           Reg.situacao_ligacao_agua = Mid(linha, 390, 20)
           Reg.situacao_ligacao_esgoto = Mid(linha, 410, 20)
           Reg.Economias_Residencial = Mid(linha, 430, 4)
           Reg.Economias_Comercial = Mid(linha, 434, 4)
           Reg.Economias_Industrial = Mid(linha, 438, 4)
           Reg.Economias_Publico = Mid(linha, 442, 4)
           Reg.Hidrometro = Mid(linha, 446, 10)
           Reg.Data_Leitura_Anterior = Mid(linha, 456, 10)
           Reg.Data_Leitura_Atual = Mid(linha, 466, 10)
           Reg.Tipo_Consumo = Mid(linha, 476, 41)
           Reg.Leitura_Anterior_Agua = Mid(linha, 517, 7)
           Reg.Constante_Consumo_Faturado_Agua = Mid(linha, 524, 8)
           Reg.Consumo_Faturado_Agua = Mid(linha, 532, 5)
           Reg.Leitura_Atual_Agua = Mid(linha, 537, 7)
           Reg.Anormalidade_Leitura_Agua = Mid(linha, 544, 25)
           Reg.Leitura_Faturada_Agua = Mid(linha, 569, 7)
           Reg.Anormalidade_Consumo_Agua = Mid(linha, 576, 25)
           Reg.Leitura_Anterior_Esgoto = Mid(linha, 601, 7)
           Reg.Constante_Consumo_Faturado_Esgoto = Mid(linha, 608, 8)
           Reg.Consumo_Faturado_Esgoto = Mid(linha, 616, 5)
           Reg.Leitura_Atual_Esgoto = Mid(linha, 621, 7)
           Reg.Anormalidade_Leitura_Esgoto = Mid(linha, 628, 25)
           Reg.Leitura_Fat_Esgoto = Mid(linha, 653, 7)
           Reg.Anormalidade_Consumo_Esgoto = Mid(linha, 660, 25)
           Reg.Consumo_Mes1 = Mid(linha, 685, 19)
           Reg.Consumo_Mes2 = Mid(linha, 704, 19)
           Reg.Consumo_Mes3 = Mid(linha, 723, 19)
           Reg.Consumo_Mes4 = Mid(linha, 742, 19)
           Reg.Consumo_Mes5 = Mid(linha, 761, 19)
           Reg.Consumo_Mes6 = Mid(linha, 780, 19)
           Reg.A1 = Mid(linha, 693, 5)
           Reg.A2 = Mid(linha, 699, 5)
           Reg.B1 = Mid(linha, 712, 5)
           Reg.B2 = Mid(linha, 718, 5)
           Reg.C1 = Mid(linha, 731, 5)
           Reg.C2 = Mid(linha, 737, 5)
           Reg.D1 = Mid(linha, 750, 5)
           Reg.D2 = Mid(linha, 756, 5)
           Reg.E1 = Mid(linha, 769, 5)
           Reg.E2 = Mid(linha, 775, 5)
           Reg.F1 = Mid(linha, 788, 5)
           Reg.F2 = Mid(linha, 794, 5)
           Reg.M1 = Mid(linha, 806, 5)
           Reg.M2 = Mid(linha, 812, 5)
           Reg.MEDIA_CONSUMO = Mid(linha, 799, 19)
           Reg.Turbidez_Exigida = Mid(linha, 818, 5)
           Reg.Turbidez_Realizada = Mid(linha, 823, 5)
           Reg.Turbidez_Conforme = Mid(linha, 828, 5)
           Reg.Cor_Aparente_Exigida = Mid(linha, 833, 5)
           Reg.Cor_Aparente_Realizada = Mid(linha, 838, 5)
           Reg.Cor_Aparente_Conforme = Mid(linha, 843, 5)
           Reg.Cloro_Residual_Livre_Exigida = Mid(linha, 848, 5)
           Reg.Cloro_Residual_Livre_Realizada = Mid(linha, 853, 5)
           Reg.Cloro_Residual_Livre_Conforme = Mid(linha, 858, 5)
           Reg.Coliformes_Totais_Exigida = Mid(linha, 863, 5)
           Reg.Coliformes_Totais_Realizada = Mid(linha, 868, 5)
           Reg.Coliformes_Totais_Conforme = Mid(linha, 873, 5)
           Reg.E_Coli_Exigida = Mid(linha, 878, 5)
           Reg.E_Coli_Realizada = Mid(linha, 883, 5)
           Reg.E_Coli_Conforme = Mid(linha, 888, 5)
           Reg.Observacao1 = Mid(linha, 893, 62)
           Reg.Observacao2 = Mid(linha, 955, 120)
           Reg.Observacao3 = Mid(linha, 1075, 100)
           Reg.Servicos1 = Mid(linha, 1175, 50)
           Reg.Consumo_Faixa1 = Mid(linha, 1225, 15)
           Reg.Valor1 = Mid(linha, 1240, 14)
           Reg.Servicos2 = Mid(linha, 1254, 50)
           Reg.Consumo_Faixa2 = Mid(linha, 1304, 15)
           Reg.Valor2 = Mid(linha, 1319, 14)
           Reg.Servicos3 = Mid(linha, 1333, 50)
           Reg.Consumo_Faixa3 = Mid(linha, 1383, 15)
           Reg.Valor3 = Mid(linha, 1398, 14)
           Reg.Servicos4 = Mid(linha, 1412, 50)
           Reg.Consumo_Faixa4 = Mid(linha, 1462, 15)
           Reg.Valor4 = Mid(linha, 1477, 14)
           Reg.Servicos5 = Mid(linha, 1491, 50)
           Reg.Consumo_Faixa5 = Mid(linha, 1541, 15)
           Reg.Valor5 = Mid(linha, 1556, 14)
           Reg.Servicos6 = Mid(linha, 1570, 50)
           Reg.Consumo_Faixa6 = Mid(linha, 1620, 15)
           Reg.Valor6 = Mid(linha, 1635, 14)
           Reg.Servicos7 = Mid(linha, 1649, 50)
           Reg.Consumo_Faixa7 = Mid(linha, 1699, 15)
           Reg.Valor7 = Mid(linha, 1714, 14)
           Reg.Servicos8 = Mid(linha, 1728, 50)
           Reg.Consumo_Faixa8 = Mid(linha, 1778, 15)
           Reg.Valor8 = Mid(linha, 1793, 14)
           Reg.Servicos9 = Mid(linha, 1807, 50)
           Reg.Consumo_Faixa9 = Mid(linha, 1857, 15)
           Reg.Valor9 = Mid(linha, 1872, 14)
           Reg.Servicos10 = Mid(linha, 1886, 50)
           Reg.Consumo_Faixa10 = Mid(linha, 1936, 15)
           Reg.Valor10 = Mid(linha, 1951, 14)
           Reg.Servicos11 = Mid(linha, 1965, 50)
           Reg.Consumo11 = Mid(linha, 2015, 15)
           Reg.Valor11 = Mid(linha, 2030, 14)
           Reg.Servicos12 = Mid(linha, 2044, 50)
           Reg.Consumo12 = Mid(linha, 2094, 15)
           Reg.Valor12 = Mid(linha, 2109, 14)
           Reg.Servicos13 = Mid(linha, 2123, 50)
           Reg.Consumo_Faixa13 = Mid(linha, 2173, 15)
           Reg.Valor13 = Mid(linha, 2188, 14)
           Reg.Servicos14 = Mid(linha, 2202, 50)
           Reg.Consumo_Faixa14 = Mid(linha, 2252, 15)
           Reg.Valor14 = Mid(linha, 2267, 14)
           Reg.Servicos15 = Mid(linha, 2281, 50)
           Reg.ConsumoFaixa15 = Mid(linha, 2331, 15)
           Reg.Valor15 = Mid(linha, 2346, 14)
           Reg.Mensagem_rateio_consumo_fixo_esgoto = Mid(linha, 2360, 31)
           Reg.PIS_Base_Calculo = Mid(linha, 2391, 14)
           Reg.PIS_Percentual = Mid(linha, 2405, 5)
           Reg.PIS_Valor_Imposto = Mid(linha, 2410, 14)
           Reg.COFINS_Base_Calculo = Mid(linha, 2424, 14)
           Reg.COFINS_Percentual = Mid(linha, 2438, 5)
           Reg.COFINS_Valor_Imposto = Mid(linha, 2443, 14)
           Reg.Data_Vencimento = Mid(linha, 2457, 10)
           Reg.Valor_Conta = Mid(linha, 2467, 14)
           Reg.Mensagem_conta1 = Mid(linha, 2481, 100)
           Reg.Mensagem_conta2 = Mid(linha, 2581, 100)
           Reg.Mensagem_conta3 = Mid(linha, 2681, 100)
           Reg.SEQUENCIAL = Mid(linha, 2781, 7)
           Reg.Código_Banco = Mid(linha, 2788, 5)
           Reg.Representacao_numerica_codigo_barras = Mid(linha, 2793, 54)
           Reg.Local_Pagamento = Mid(linha, 2847, 45)
           Reg.VENCIMENTO = Mid(linha, 2892, 10)
           Reg.Cedente = Mid(linha, 2912, 50)
           Reg.Agencia_Cedente = Mid(linha, 2962, 13)
           Reg.Data_Documento = Mid(linha, 2975, 10)
           Reg.Numero_Documento = Mid(linha, 2985, 10)
           Reg.Especie_Documento = Mid(linha, 2995, 3)
           Reg.Aceite = Mid(linha, 2998, 1)
           Reg.Data_Processamento = Mid(linha, 2999, 10)
           Reg.Nosso_Numero = Mid(linha, 3009, 19)
           Reg.Carteira = Mid(linha, 3028, 2)
           Reg.Valor_Documento = Mid(linha, 3030, 14)
           Reg.Sacado_linha_1a = Mid(linha, 3044, 50)
           Reg.Sacado_linha1b = Mid(linha, 3094, 16)
           Reg.Sacado_linha1c = Mid(linha, 3110, 10)
           Reg.Sacado_linha1d = Mid(linha, 3120, 14)
           Reg.Sacado_linha1e = Mid(linha, 3134, 9)
           Reg.Sacado_linha2a = Mid(linha, 3143, 50)
           Reg.Sacado_linha3a = Mid(linha, 3193, 50)
           Reg.Codigo_Barras = Mid(linha, 3243, 112)
           Reg.CPF_CNPJ = Trim(Mid(linha, 3355, 18))
           Reg.MSG01 = Trim(Mid(linha, 3373, 42))
           Reg.MSG02 = Trim(Mid(linha, 3415, 30))
           Reg.MSG03 = Trim(Mid(linha, 3445, 143))
           Reg.MSG04 = Trim(Mid(linha, 3588, 61))
           
           If Len(Trim(Reg.NOME_CONTA)) > 0 Then
              MontaExtrato
              LimpaReg
              
              totRegs = totRegs + 1
              frmPrincipal.txQtdReg.Text = totRegs
           End If

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
        Print #outFN, "/ImageData2"
        Print #outFN, "currentfile"
        Print #outFN, "<< /Filter /SubFileDecode"
        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
        Print #outFN, ">> /ReusableStreamDecode filter"
        Open App.Path & "\Conta_Boleto_A4_Ext.ps" For Input As #200
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
        'Terceira Imagem
        Print #outFN, "/ImageData3"
        Print #outFN, "currentfile"
        Print #outFN, "<< /Filter /SubFileDecode"
        Print #outFN, "   /DecodeParms << /EODCount 0 /EODString (*EOD*) >>"
        Print #outFN, ">> /ReusableStreamDecode filter"
        Open App.Path & "\Conta_Boleto_A4_Int_Resp.ps" For Input As #200
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
           Reg.MENSAGEM = " "
           Reg.Nome_Localidade = " "
           Reg.Matrícula_Imovel = " "
           Reg.MES_ANO = " "
           Reg.NOME_CONTA = " "
           Reg.ENDERECO_IMOVEL = " "
           Reg.Endereco_Imovel2 = " "
           Reg.INSCRICAO = " "
           Reg.Grupo_Faturamento = " "
           Reg.Codigo_debito_automatico = " "
           Reg.Codigo_cliente_responsável = " "
           Reg.Endereco_correspondencia_responsavel = " "
           Reg.Endereco_correspondencia_responsavel2 = " "
           Reg.situacao_ligacao_agua = " "
           Reg.situacao_ligacao_esgoto = " "
           Reg.Economias_Residencial = " "
           Reg.Economias_Comercial = " "
           Reg.Economias_Industrial = " "
           Reg.Economias_Publico = " "
           Reg.Hidrometro = " "
           Reg.Data_Leitura_Anterior = " "
           Reg.Data_Leitura_Atual = " "
           Reg.Tipo_Consumo = " "
           Reg.Leitura_Anterior_Agua = " "
           Reg.Constante_Consumo_Faturado_Agua = " "
           Reg.Consumo_Faturado_Agua = " "
           Reg.Leitura_Atual_Agua = " "
           Reg.Anormalidade_Leitura_Agua = " "
           Reg.Leitura_Faturada_Agua = " "
           Reg.Anormalidade_Consumo_Agua = " "
           Reg.Leitura_Anterior_Esgoto = " "
           Reg.Constante_Consumo_Faturado_Esgoto = " "
           Reg.Consumo_Faturado_Esgoto = " "
           Reg.Leitura_Atual_Esgoto = " "
           Reg.Anormalidade_Leitura_Esgoto = " "
           Reg.Leitura_Fat_Esgoto = " "
           Reg.Anormalidade_Consumo_Esgoto = " "
           Reg.Consumo_Mes1 = " "
           Reg.Consumo_Mes2 = " "
           Reg.Consumo_Mes3 = " "
           Reg.Consumo_Mes4 = " "
           Reg.Consumo_Mes5 = " "
           Reg.Consumo_Mes6 = " "
           Reg.MEDIA_CONSUMO = " "
           Reg.Turbidez_Exigida = " "
           Reg.Turbidez_Realizada = " "
           Reg.Turbidez_Conforme = " "
           Reg.Cor_Aparente_Exigida = " "
           Reg.Cor_Aparente_Realizada = " "
           Reg.Cor_Aparente_Conforme = " "
           Reg.Cloro_Residual_Livre_Exigida = " "
           Reg.Cloro_Residual_Livre_Realizada = " "
           Reg.Cloro_Residual_Livre_Conforme = " "
           Reg.Coliformes_Totais_Exigida = " "
           Reg.Coliformes_Totais_Realizada = " "
           Reg.Coliformes_Totais_Conforme = " "
           Reg.E_Coli_Exigida = " "
           Reg.E_Coli_Realizada = " "
           Reg.E_Coli_Conforme = " "
           Reg.Observacao1 = " "
           Reg.Observacao2 = " "
           Reg.Observacao3 = " "
           Reg.Servicos1 = " "
           Reg.Consumo_Faixa1 = " "
           Reg.Valor1 = " "
           Reg.Servicos2 = " "
           Reg.Consumo_Faixa2 = " "
           Reg.Valor2 = " "
           Reg.Servicos3 = " "
           Reg.Consumo_Faixa3 = " "
           Reg.Valor3 = " "
           Reg.Servicos4 = " "
           Reg.Consumo_Faixa4 = " "
           Reg.Valor4 = " "
           Reg.Servicos5 = " "
           Reg.Consumo_Faixa5 = " "
           Reg.Valor5 = " "
           Reg.Servicos6 = " "
           Reg.Consumo_Faixa6 = " "
           Reg.Valor6 = " "
           Reg.Servicos7 = " "
           Reg.Consumo_Faixa7 = " "
           Reg.Valor7 = " "
           Reg.Servicos8 = " "
           Reg.Consumo_Faixa8 = " "
           Reg.Valor8 = " "
           Reg.Servicos9 = " "
           Reg.Consumo_Faixa9 = " "
           Reg.Valor9 = " "
           Reg.Servicos10 = " "
           Reg.Consumo_Faixa10 = " "
           Reg.Valor10 = " "
           Reg.Servicos11 = " "
           Reg.Consumo11 = " "
           Reg.Valor11 = " "
           Reg.Servicos12 = " "
           Reg.Consumo12 = " "
           Reg.Valor12 = " "
           Reg.Servicos13 = " "
           Reg.Consumo_Faixa13 = " "
           Reg.Valor13 = " "
           Reg.Servicos14 = " "
           Reg.Consumo_Faixa14 = " "
           Reg.Valor14 = " "
           Reg.Servicos15 = " "
           Reg.ConsumoFaixa15 = " "
           Reg.Valor15 = " "
           Reg.Mensagem_rateio_consumo_fixo_esgoto = " "
           Reg.PIS_Base_Calculo = " "
           Reg.PIS_Percentual = " "
           Reg.PIS_Valor_Imposto = " "
           Reg.COFINS_Base_Calculo = " "
           Reg.COFINS_Percentual = " "
           Reg.COFINS_Valor_Imposto = " "
           Reg.Data_Vencimento = " "
           Reg.Valor_Conta = " "
           Reg.Mensagem_conta1 = " "
           Reg.Mensagem_conta2 = " "
           Reg.Mensagem_conta3 = " "
           Reg.SEQUENCIAL = " "
           Reg.Código_Banco = " "
           Reg.Representacao_numerica_codigo_barras = " "
           Reg.Local_Pagamento = " "
           Reg.VENCIMENTO = " "
           Reg.Cedente = " "
           Reg.Agencia_Cedente = " "
           Reg.Data_Documento = " "
           Reg.Numero_Documento = " "
           Reg.Especie_Documento = " "
           Reg.Aceite = " "
           Reg.Data_Processamento = " "
           Reg.Nosso_Numero = " "
           Reg.Carteira = " "
           Reg.Valor_Documento = " "
           Reg.Sacado_linha_1a = " "
           Reg.Sacado_linha1b = " "
           Reg.Sacado_linha1c = " "
           Reg.Sacado_linha1d = " "
           Reg.Sacado_linha1e = " "
           Reg.Sacado_linha2a = " "
           Reg.Sacado_linha3a = " "
           Reg.Codigo_Barras = " "

End Sub


