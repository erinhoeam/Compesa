Attribute VB_Name = "mGeneral"
'*****************************************************
'Desenvolvido por Aziel R. Pereira Jr
'                 Analista de Sistemas
'Data: Junho/2004
'Contatos - Email: azieljr@hotmail.com
'           Fone: (81) 9979.1972
'Aplicação: Módulo de Funções Diversas
'*****************************************************

Option Explicit

Public Const A3 = 103
Public Const A4 = 104
Public Const DUPLOA4 = 204
Public wStringConexao As String

'Faz a formatação dos números, já que a função do Visual Basic não consegue colocar a _
vírgula no lugar certo quando é passada uma string com números. _
Ex.: Ao ser utilizado Format("1234567890","###,###,###.##"), o resultado é: "1.234.567.890," _
     ou Format("1234567890","###.###.###,##"), o resultado é: "1234567890,."
Public Function convVlr(vlr As String, comZeros As Boolean) As String
   Dim i As Integer
   Dim tamStr As Integer
   Dim countPoint As Integer
   
   'Testar se não é tudo branco, se for nao fazer processamento
   convVlr = vlr
   vlr = Trim(vlr)
   
   If vlr = "" Then
      'Antes de sair da função, testar se é necessário colocar zeros ou não
      If comZeros Then _
         convVlr = "0,00"
      Exit Function
   End If
   
   tamStr = Len(vlr) 'Tamanho do valor passado
   i = 1
   'Enquanto for "0", preencher com " " (espaços em branco) (Fazer até tamanho - 2 para não _
    limpar os decimais, que podem ser 0
   While Mid(vlr, i, 1) = "0" And i < tamStr - 2
      Mid(vlr, i, 1) = " "
      i = i + 1
   Wend

   'Iniciar o resultado com os decimais e a vírgula ",xx"
   convVlr = "," & Right(vlr, 2)
   
   i = Len(vlr) - 2     'Tamanho total - 2
   countPoint = 0
   While i > 0
      convVlr = Mid(vlr, i, 1) & convVlr  'Pegar cada número, de trás pra frente, e colocar _
                                           na variável de saída, na frente dos decimais
      countPoint = countPoint + 1
      If i > 1 Then
         'Se o próximo valor a ser lido for um número e se o contador for 3, então colocar _
          o ponto dos decimais
         If Mid(vlr, i - 1, 1) <> " " And countPoint = 3 Then
            convVlr = "." & convVlr
            countPoint = 0
         End If
      End If
      i = i - 1   'Decremento do while
   Wend
End Function

'Ler o arquivo byte a byte. _
Motivo: Alguns arquivos enviados pelo hiper chegam no _
formato UNIX (apenas LF no final da linha) e a função "line input" do Visual Basic _
só reconhece o CRLF como final de linha.
Public Function convVlr5(vlr As String, comZeros As Boolean) As String
   Dim i As Integer
   Dim tamStr As Integer
   Dim countPoint As Integer
   
   'Testar se não é tudo branco, se for nao fazer processamento
   convVlr5 = vlr
   vlr = Trim(vlr)
   
   If vlr = "" Then
      'Antes de sair da função, testar se é necessário colocar zeros ou não
      If comZeros Then _
         convVlr5 = "0,00000"
      Exit Function
   End If
   
   tamStr = Len(vlr) 'Tamanho do valor passado
   i = 1
   'Enquanto for "0", preencher com " " (espaços em branco) (Fazer até tamanho - 2 para não _
    limpar os decimais, que podem ser 0
   While Mid(vlr, i, 1) = "0" And i < tamStr - 5
      Mid(vlr, i, 1) = " "
      i = i + 1
   Wend

   'Iniciar o resultado com os decimais e a vírgula ",xx"
   convVlr5 = "," & Right(vlr, 5)
   
   i = Len(vlr) - 5     'Tamanho total - 2
   countPoint = 0
   While i > 0
      convVlr5 = Mid(vlr, i, 1) & convVlr5  'Pegar cada número, de trás pra frente, e colocar _
                                             na variável de saída, na frente dos decimais
      countPoint = countPoint + 1
      If i > 1 Then
         'Se o próximo valor a ser lido for um número e se o contador for 3, então colocar _
          o ponto dos decimais
         If Mid(vlr, i - 1, 1) <> " " And countPoint = 3 Then
            convVlr5 = "." & convVlr5
            countPoint = 0
         End If
      End If
      i = i - 1   'Decremento do while
   Wend
End Function

'Ler o arquivo byte a byte. _
Motivo: Alguns arquivos enviados pelo hiper chegam no _
formato UNIX (apenas LF no final da linha) e a função "line input" do Visual Basic _
só reconhece o CRLF como final de linha.
Public Function convVlr2(vlr As String, comZeros As Boolean) As String
   Dim i As Integer
   Dim tamStr As Integer
   Dim countPoint As Integer
   
   'Testar se não é tudo branco, se for nao fazer processamento
   convVlr2 = vlr
   vlr = Trim(vlr)
   
   If vlr = "" Then
      'Antes de sair da função, testar se é necessário colocar zeros ou não
      If comZeros Then _
         convVlr2 = "R$ 0,00"
      Exit Function
   End If
   
   tamStr = Len(vlr) 'Tamanho do valor passado
   i = 1
   'Enquanto for "0", preencher com " " (espaços em branco) (Fazer até tamanho - 2 para não _
    limpar os decimais, que podem ser 0
   While Mid(vlr, i, 1) = "0" And i < tamStr - 2
      Mid(vlr, i, 1) = " "
      i = i + 1
   Wend

   'Iniciar o resultado com os decimais e a vírgula ",xx"
   convVlr2 = "," & Right(vlr, 2)
   
   i = Len(vlr) - 2     'Tamanho total - 2
   countPoint = 0
   While i > 0
      convVlr2 = Mid(vlr, i, 1) & convVlr2  'Pegar cada número, de trás pra frente, e colocar _
                                           na variável de saída, na frente dos decimais
      countPoint = countPoint + 1
      If i > 1 Then
         'Se o próximo valor a ser lido for um número e se o contador for 3, então colocar _
          o ponto dos decimais
         If Mid(vlr, i - 1, 1) <> " " And countPoint = 3 Then
            convVlr2 = "." & convVlr2
            countPoint = 0
         End If
      End If
      i = i - 1   'Decremento do while
   Wend
   convVlr2 = "R$ " & Trim(convVlr2)
End Function
Public Sub ler(fileNumber As Integer, str As String)
   Dim char As Byte
   'Limpar a string de saída
   str = ""
   'Testar se não for fim de arquivo
   If Not EOF(fileNumber) Then
      'Pegar um caracter do arquivo
      Get #fileNumber, , char
      'Continuar pegando enquanto não for LF ou CR ou EOF
      Do While char <> 10 And char <> 0
         'Concatenar na string cada caracter lido
          If char <> 13 Then
             str = str & Chr(char)
          End If
         'Pegar um caracter do arquivo
         Get #fileNumber, , char
      Loop
      'Se sair do Loop por causa do CR, pegar o próximo caracter (LF)
      If char = 13 Then _
         Get #fileNumber, , char
   End If
End Sub

'Para exibir os caracteres "\", "(" e ")" no VIPP/Postscript é necessário colocar um "\" antes, _
o qual é o motivo dessa função. Caso isto não seja feito, podem ocorrer erros de processamento _
do VIPP/Postscript pela impressora.
Public Function convStr(str As String) As String
   Dim testChar As Integer
   'Procurar a posição do primeiro "\" na string
   testChar = InStr(1, str, "\", vbBinaryCompare)
   'Caso exista algum, sua posição será diferente de "0"
   While testChar > 0
      'Colocar um "\" antes do "\" encontrado
      str = Left$(str, testChar - 1) & "\" & Mid(str, testChar, Len(str))
      'Pular duas posições, já que foi colocado um "\" na posição encontrada, empurrando _
       o resto dos caracteres em 1
      testChar = testChar + 2
      'Testar a partir da nova posição
      testChar = InStr(testChar, str, "\", vbBinaryCompare)
   Wend

'Fazer o mesmo para "(" e ")"

   testChar = InStr(1, str, "(", vbBinaryCompare)
   While testChar > 0
      str = Left$(str, testChar - 1) & "\" & Mid(str, testChar, Len(str))
      testChar = testChar + 2
      testChar = InStr(testChar, str, "(", vbBinaryCompare)
   Wend
   
   testChar = InStr(1, str, ")", vbBinaryCompare)
   While testChar > 0
      str = Left$(str, testChar - 1) & "\" & Mid(str, testChar, Len(str))
      testChar = testChar + 2
      testChar = InStr(testChar, str, ")", vbBinaryCompare)
   Wend
   
   'Procura por caracteres decimais #23 e #26 e troca por espaço em branco
   'Foram detectados estes caracteres nos dados das Fichas de Cobrança reg1030-nprof 20.11.2001
   testChar = InStr(1, str, Chr(23), vbBinaryCompare)
   While testChar > 0
      str = Left$(str, testChar - 1) & " " & Mid(str, testChar + 1, Len(str))
      testChar = InStr(testChar, str, Chr(23), vbBinaryCompare)
   Wend
   testChar = InStr(1, str, Chr(26), vbBinaryCompare)
   While testChar > 0
      str = Left$(str, testChar - 1) & " " & Mid(str, testChar + 1, Len(str))
      testChar = InStr(testChar, str, Chr(26), vbBinaryCompare)
   Wend
   
   convStr = str
End Function

Public Function convData(dt As String) As String
   convData = Right(dt, 2) & "/" & Mid(dt, 5, 2) & "/" & Left(dt, 4)
End Function

Public Function convDataPorExt(dt As String) As String
Dim mes(12) As String * 20
   mes(0) = ""
   mes(1) = "janeiro"
   mes(2) = "feveiro"
   mes(3) = "março"
   mes(4) = "abril"
   mes(5) = "maio"
   mes(6) = "junho"
   mes(7) = "julho"
   mes(8) = "agosto"
   mes(9) = "setembro"
   mes(10) = "outubro"
   mes(11) = "novembro"
   mes(12) = "dezembro"
   convDataPorExt = Format(Right(dt, 2), "##") & " de " & Trim(mes(CInt(Mid(dt, 5, 2)))) & " de " & Left(dt, 4)
End Function


Public Function nomeDoJob(job As String) As String
   Dim iJob As Integer, jJob As Integer
   
   iJob = Len(job)
   jJob = 0
   Do While Mid(job, iJob, 1) <> "\"
      iJob = iJob - 1
      jJob = jJob + 1
   Loop
   nomeDoJob = Right(job, jJob)
End Function

Public Function convCPFCGC(num As String) As String
'00.000.000/0000-00
'000.000.000-00

   num = Trim(num)
   
   convCPFCGC = num
   
   If num <> "" Then
      If Len(num) = 14 Then
         convCPFCGC = Left$(num, 2) & "." & Mid(num, 3, 3) & "." & Mid(num, 6, 3) & "/" & Mid(num, 9, 4) & "-" & Mid(num, 13, 2)
      Else
         convCPFCGC = Left(num, 3) & "." & Mid(num, 4, 3) & "." & Mid(num, 7, 3) & "-" & Mid(num, 10, 2) & "    "
      End If
   End If
End Function

Public Sub GeraArqIndices(iFile As String, pProc As String, pGer As String, pRecSize As Integer, ByRef qtdCli As Long)
   Dim jobName As String
   Dim oFile As String
   Dim inFN As Integer, indFN As Integer, numlinha As Long, numfac
   Dim linha As String * 141

   qtdCli = 0
   numlinha = 0
   numfac = 1
   inFN = FreeFile
   Open iFile For Random As #inFN Len = pRecSize
   
   jobName = nomeDoJob(iFile)
   indFN = FreeFile
   oFile = pGer & Left$(jobName, Len(jobName) - 4) & ".ind"
   Open oFile For Output As indFN
   
   While Not EOF(inFN)
      Get #inFN, , linha
      numlinha = numlinha + 1
      If Left(linha, 4) = "1001" Then
         Print #indFN, Format(numlinha, "00000000") & " " & Format(numfac, "0000")
         qtdCli = qtdCli + 1
      ElseIf Left(linha, 4) = "1500" Then
         numfac = numfac + 1
      End If
   Wend
   
   Close inFN
   Close indFN
End Sub

Public Sub OrdenaArquivoTwoUp(iFile As String, pGer As String)
'Usado na ordenação das Fichas de Cobrança
   Dim qtdCli As Long
   Dim jobName As String
   Dim indFile As String, oFile As String
   Dim inFN As Integer, indFN As Integer, gerFN As Integer, numlinha As Long, numfac
   Dim linha As String
   Dim IniInt1 As Long, IniInt2 As Long, IniExt1 As Long, IniExt2 As Long
   Dim cont1 As Long, cont2 As Long, cont3 As Long, cont4 As Long
   Dim Header As Boolean
   Dim linha148 As String * 148
   Dim linha27 As String * 27
   Dim i As Long

   qtdCli = 0
   numlinha = 0
   inFN = FreeFile
   Open iFile For Input As #inFN
    
   jobName = nomeDoJob(iFile)
   indFN = FreeFile
   oFile = pGer & Left$(jobName, Len(jobName) - 4) & ".ind"
   Open oFile For Output As indFN
   gerFN = FreeFile
   oFile = pGer & Left$(jobName, Len(jobName) - 4) & ".tmp2"
   Open oFile For Output As gerFN
   
   Header = True
   While Header
      Input #inFN, linha
      linha148 = linha
      numlinha = numlinha + 1
      Print #gerFN, linha148
      If Left(linha, 13) = "%% End Header" Then Header = False
   Wend
   
   IniInt1 = 0
   IniInt2 = 0
   IniExt1 = 0
   IniExt2 = 0
   
   cont1 = 0
   cont2 = 0
   cont3 = 0
   cont4 = 0
  
   While Not EOF(inFN)
      Line Input #inFN, linha
      numlinha = numlinha + 1
      linha148 = linha
      Print #gerFN, linha148
      'Debug.Print numlinha, linha
      If Left(linha, 8) = "% IniInt" Then
         If IniInt1 = 0 Then
            IniInt1 = numlinha
         Else
            IniInt2 = numlinha
         End If
         qtdCli = qtdCli + 1
      ElseIf Left(linha, 8) = "% IniExt" Then
         If IniExt1 = 0 Then
            IniExt1 = numlinha
         Else
            IniExt2 = numlinha
         End If
      ElseIf Left(linha, 8) = "% FimInt" Then
         If cont1 = 0 Then
            cont1 = numlinha - IniInt1 + 1
         ElseIf cont2 = 0 Then
            cont2 = numlinha - IniInt2 + 1
         End If
      ElseIf Left(linha, 8) = "% FimExt" Then
         If cont3 = 0 Then
            cont3 = numlinha - IniExt1 + 1
         ElseIf cont4 = 0 Then
            cont4 = numlinha - IniExt2 + 1
            Print #indFN, Format(IniInt1, "00000000") & " " & Format(cont1, "0000") & " " & Format(IniExt1, "00000000") & " " & Format(cont3, "0000")
            Print #indFN, Format(IniInt2, "00000000") & " " & Format(cont2, "0000") & " " & Format(IniExt2, "00000000") & " " & Format(cont4, "0000")
            IniInt1 = 0
            IniInt2 = 0
            IniExt1 = 0
            IniExt2 = 0
            cont1 = 0
            cont2 = 0
            cont3 = 0
            cont4 = 0
         End If
      End If
   Wend
 
   Close inFN
   Close indFN
   Close gerFN
   
   'Arquivo gerado .lst de tamanho fixo 150 bytes
   inFN = FreeFile
   iFile = pGer & Left$(jobName, Len(jobName) - 4) & ".tmp2"
   Open iFile For Random As #inFN Len = 150
   'Arquivo de índices de tamanho fixo 29 bytes
   indFN = FreeFile
   indFile = pGer & Left$(jobName, Len(jobName) - 4) & ".ind"
   Open indFile For Random As indFN Len = 29
   'Arquivo a ser gerado ordenado
   gerFN = FreeFile
   oFile = pGer & Left$(jobName, Len(jobName) - 4) & ".lst"
   Open oFile For Output As gerFN
   
   Header = True
   While Header
      Get #inFN, , linha148
      Print #gerFN, Trim(linha148)
      If Left(linha148, 13) = "%% End Header" Then Header = False
   Wend
   
   'Formato do registro de indice
   'Ini-Int  Qtd  Ini-Ext  Qtd
   'AAAAAAAA BBBB CCCCCCCC DDDD
   
   'frmPrincipal.ProgressBarProcess.Max = qtdCli / 2
   'frmPrincipal.ProgressBarProcess.Value = 0.0001
   
   
   For i = 1 To Int(qtdCli / 2)
      'If i <= frmPrincipal.ProgressBarProcess.Max Then frmPrincipal.ProgressBarProcess.Value = i
   
   'Imprime linhas em branco
      Print #gerFN,
      Print #gerFN,
   'Imprima o Interno do primeiro bloco %IniInt-01
      Get #indFN, i, linha27
      IniInt1 = Val(Mid(linha27, 1, 8))
      Get #inFN, IniInt1, linha148
      For cont1 = 1 To Val(Mid(linha27, 10, 4))
         Get #inFN, IniInt1 + cont1 - 1, linha148
         Print #gerFN, Trim(linha148)
      Next
   
   'Imprima o Interno do Segundo bloco %IniInt-02
      Get #indFN, i + Int(qtdCli / 2), linha27
      IniInt2 = Val(Mid(linha27, 1, 8))
      Get #inFN, IniInt2, linha148
      For cont2 = 1 To Val(Mid(linha27, 10, 4))
         Get #inFN, IniInt2 + cont2 - 1, linha148
         Print #gerFN, Trim(linha148)
      Next
   
      'Imprima o Externo do primeiro bloco %IniExt-01
      Get #indFN, i, linha27
      IniInt1 = Val(Mid(linha27, 15, 8))
      Get #inFN, IniInt1, linha148
      For cont1 = 1 To Val(Mid(linha27, 24, 4))
         Get #inFN, IniInt1 + cont1 - 1, linha148
         Print #gerFN, Trim(linha148)
      Next
      
      'Imprima o Externo do segundo bloco %IniExt-02
      Get #indFN, i + Int(qtdCli / 2), linha27
      IniInt2 = Val(Mid(linha27, 15, 8))
      Get #inFN, IniInt2, linha148
      For cont2 = 1 To Val(Mid(linha27, 24, 4))
         Get #inFN, IniInt2 + cont2 - 1, linha148
         Print #gerFN, Trim(linha148)
      Next
   
   Next
   
   Close inFN
   Close indFN
   Close gerFN
   
   'Elimina os arquivos auxiliares utilizados
   If Manip_Arq(EXISTEARQ, iFile) Then
      If Not Manip_Arq(DELETAR, iFile) Then
         MsgBox "O arquivo " & iFile & " não pode ser excluído.", vbCritical
      End If
   End If
   
   iFile = pGer & Left$(jobName, Len(jobName) - 4) & ".tmp"
   If Manip_Arq(EXISTEARQ, iFile) Then
      If Not Manip_Arq(DELETAR, iFile) Then
         MsgBox "O arquivo " & iFile & " não pode ser excluído.", vbCritical
      End If
   End If
   
   If Manip_Arq(EXISTEARQ, indFile) Then
      If Not Manip_Arq(DELETAR, indFile) Then
         MsgBox "O arquivo " & indFile & " não pode ser excluído.", vbCritical
      End If
   End If
   
   'frmPrincipal.ProgressBarProcess.Value = 0.0001
   
End Sub
Public Function cmDot(Dot3 As Integer) As Double
       cmDot = Int(Dot3 * 118.11)
End Function
Public Function cm(cor As Double) As Double
       cm = Int(cor * 567)
End Function

Function extenso(VALOR As Double)
Dim wtexto1, wtexto2, wvalor As Double
Dim n00, n01, n02, n03, n04, n05, n06, n07, n08, n09, n10, n11, n12, n13, n14, n15 As String
Dim n16, n17, n18, n19, n20, n30, n40, n50, n60, n70, n80, n90, n000, n100, n100a As String
Dim n200, n300, n400, n500, n600, n700, n800, n900, mil, milhao, milhoes, moeda, moedas, centavo, centavos As String
Dim Start, wInteiro, wDecimal, dec1dig As String
Dim vlr1dig, vlr2dig, vlr3dig, vlr4dig, vlr5dig, vlr6dig, vlr7dig, vlr8dig, dec2dig As String

wtexto1 = "": wtexto2 = "": wvalor = VALOR
' Criar as variaveis de valores
n01 = "UM ": n02 = "DOIS ": n03 = "TRÊS ": n04 = "QUATRO "
n05 = "CINCO ": n06 = "SEIS ": n07 = "SETE ": n08 = "OITO "
n09 = "NOVE ": n00 = " ": n10 = "DEZ ": n11 = "ONZE "
n12 = "DOZE ": n13 = "TREZE ": n14 = "QUATORZE ": n15 = "QUINZE "
n16 = "DEZESEIS ": n17 = "DEZESETE ": n18 = "DEZOITO ": n19 = "DEZENOVE "
n20 = "VINTE ": n30 = "TRINTA ": n40 = "QUARENTA ": n50 = "CINQUENTA "
n60 = "SESSENTA ": n70 = "SETENTA ": n80 = "OITENTA ": n90 = "NOVENTA "
n000 = " ": n100 = "CEM ": n100a = "CENTO ": n200 = "DUZENTOS ": n300 = "TREZENTOS "
n400 = "QUATROCENTOS ": n500 = "QUINHENTOS ": n600 = "SEISCENTOS ": n700 = "SETECENTOS "
n800 = "OITOCENTOS ": n900 = "NOVECENTOS ": mil = "MIL ": milhao = "MILHAO "
milhoes = "MILHOES ": moeda = "REAL ": moedas = "REAIS ": centavo = "CENTAVO ": centavos = "CENTAVOS "

' Define as casas inteiras e decimais
Start = InStr(1, wvalor, ",")
If Start = 0 Then
   wInteiro = Trim(wvalor):   wDecimal = "00"
Else
   wInteiro = Mid(wvalor, 1, InStr(1, wvalor, ",") - 1)
   wDecimal = Mid(wvalor, InStr(1, wvalor, ",") + 1, 2)
   If Len(Trim(wDecimal)) = 1 Then  ' Teste para quando a segunda
      wDecimal = wDecimal + "0"     ' casa do centavo for igual a zero
   End If
End If
' Define onde começa a escrever o extenso
If Len(Trim(wInteiro)) = 1 Then
   vlr1dig = wInteiro:   GoSub rot1dig
End If
If Len(Trim(wInteiro)) = 2 Then
   vlr2dig = wInteiro:   GoSub rot2dig
End If
If Len(Trim(wInteiro)) = 3 Then
   vlr3dig = wInteiro:   GoSub rot3dig
End If
If Len(Trim(wInteiro)) = 4 Then
   vlr4dig = wInteiro:   GoSub rot4dig
End If
If Len(Trim(wInteiro)) = 5 Then
   vlr5dig = wInteiro:   GoSub rot5dig
End If
If Len(Trim(wInteiro)) = 6 Then
   vlr6dig = wInteiro:   GoSub rot6dig
End If
If Len(Trim(wInteiro)) = 7 Then
   vlr7dig = wInteiro:   GoSub rot7dig
End If
If Len(Trim(wInteiro)) = 8 Then
   vlr8dig = wInteiro:   GoSub rot8dig
End If
' Gerar o extenso do decimal
If Len(Trim(wDecimal)) > 0 Then
   If wDecimal <> "00" Then
      dec2dig = wDecimal: GoSub rot2digdec
   End If
End If
If Len(Trim(wInteiro)) > 8 Then
   extenso = " Esta rotina nao suporta mais de 8 numeros inteiros "
   Exit Function
End If

' Finaliza e retorna a string
extenso = wtexto1 & wtexto2
Exit Function
' ----------------------------------------
' Rotina para um digito inteiro de extenso
rot1dig:
If vlr1dig = "1" Then
   If Val(vlr1dig) = 1 Then
      wtexto1 = wtexto1 & n01 & moeda
   Else
     wtexto1 = wtexto1 & n01 & moedas
   End If
ElseIf vlr1dig = "2" Then
   wtexto1 = wtexto1 & n02 & moedas
ElseIf vlr1dig = "3" Then
   wtexto1 = wtexto1 & n03 & moedas
ElseIf vlr1dig = "4" Then
   wtexto1 = wtexto1 & n04 & moedas
ElseIf vlr1dig = "5" Then
   wtexto1 = wtexto1 & n05 & moedas
ElseIf vlr1dig = "6" Then
   wtexto1 = wtexto1 & n06 & moedas
ElseIf vlr1dig = "7" Then
   wtexto1 = wtexto1 & n07 & moedas
ElseIf vlr1dig = "8" Then
   wtexto1 = wtexto1 & n08 & moedas
ElseIf vlr1dig = "9" Then
   wtexto1 = wtexto1 & n09 & moedas
End If
Return

' Rotina para dois digitos inteiros de extenso
rot2dig:
If Mid(vlr2dig, 1, 1) = "0" Then
   vlr1dig = Mid(vlr2dig, 2, 1)
   GoSub rot1dig
   Return
End If

If Val(vlr2dig) < 20 Then
   If vlr2dig = "10" Then wtexto1 = wtexto1 & n10 & moedas
   If vlr2dig = "11" Then wtexto1 = wtexto1 & n11 & moedas
   If vlr2dig = "12" Then wtexto1 = wtexto1 & n12 & moedas
   If vlr2dig = "13" Then wtexto1 = wtexto1 & n13 & moedas
   If vlr2dig = "14" Then wtexto1 = wtexto1 & n14 & moedas
   If vlr2dig = "15" Then wtexto1 = wtexto1 & n15 & moedas
   If vlr2dig = "16" Then wtexto1 = wtexto1 & n16 & moedas
   If vlr2dig = "17" Then wtexto1 = wtexto1 & n17 & moedas
   If vlr2dig = "18" Then wtexto1 = wtexto1 & n18 & moedas
   If vlr2dig = "19" Then wtexto1 = wtexto1 & n19 & moedas
   Return
Else
   If Mid(vlr2dig, 1, 1) = "2" Then wtexto1 = wtexto1 & n20
   If Mid(vlr2dig, 1, 1) = "3" Then wtexto1 = wtexto1 & n30
   If Mid(vlr2dig, 1, 1) = "4" Then wtexto1 = wtexto1 & n40
   If Mid(vlr2dig, 1, 1) = "5" Then wtexto1 = wtexto1 & n50
   If Mid(vlr2dig, 1, 1) = "6" Then wtexto1 = wtexto1 & n60
   If Mid(vlr2dig, 1, 1) = "7" Then wtexto1 = wtexto1 & n70
   If Mid(vlr2dig, 1, 1) = "8" Then wtexto1 = wtexto1 & n80
   If Mid(vlr2dig, 1, 1) = "9" Then wtexto1 = wtexto1 & n90
End If
If Mid(vlr2dig, 2, 1) = "0" Then
   wtexto1 = wtexto1 & moedas
Else
   wtexto1 = wtexto1 & " E "
   vlr1dig = Mid(vlr2dig, 2, 1)
   GoSub rot1dig
End If
Return

' Rotina para treis digitos inteiros de extenso
rot3dig:
If Mid(vlr3dig, 1, 1) = "0" Then
   vlr2dig = Mid(vlr3dig, 2, 2)
   GoSub rot2dig
   Return
End If

If vlr3dig = "100" Then
   wtexto1 = wtexto1 & n100 & moedas
   Return
End If
   
If Mid(vlr3dig, 1, 1) = "1" Then wtexto1 = wtexto1 & n100a
If Mid(vlr3dig, 1, 1) = "2" Then wtexto1 = wtexto1 & n200
If Mid(vlr3dig, 1, 1) = "3" Then wtexto1 = wtexto1 & n300
If Mid(vlr3dig, 1, 1) = "4" Then wtexto1 = wtexto1 & n400
If Mid(vlr3dig, 1, 1) = "5" Then wtexto1 = wtexto1 & n500
If Mid(vlr3dig, 1, 1) = "6" Then wtexto1 = wtexto1 & n600
If Mid(vlr3dig, 1, 1) = "7" Then wtexto1 = wtexto1 & n700
If Mid(vlr3dig, 1, 1) = "8" Then wtexto1 = wtexto1 & n800
If Mid(vlr3dig, 1, 1) = "9" Then wtexto1 = wtexto1 & n900

If Mid(vlr3dig, 2, 2) = "00" Then
   wtexto1 = wtexto1 & moedas
Else
   wtexto1 = wtexto1 & " E "
   vlr2dig = Mid(vlr3dig, 2, 2)
   GoSub rot2dig
End If
Return

' Rotina para Quatro digitos inteiros de extenso
rot4dig:
If Mid(vlr4dig, 1, 1) = "0" Then
   vlr3dig = Mid(vlr4dig, 2, 3)
   GoSub rot3dig
   Return
End If
If Mid(vlr4dig, 1, 1) = "1" Then wtexto1 = wtexto1 & n01 & mil
If Mid(vlr4dig, 1, 1) = "2" Then wtexto1 = wtexto1 & n02 & mil
If Mid(vlr4dig, 1, 1) = "3" Then wtexto1 = wtexto1 & n03 & mil
If Mid(vlr4dig, 1, 1) = "4" Then wtexto1 = wtexto1 & n04 & mil
If Mid(vlr4dig, 1, 1) = "5" Then wtexto1 = wtexto1 & n05 & mil
If Mid(vlr4dig, 1, 1) = "6" Then wtexto1 = wtexto1 & n06 & mil
If Mid(vlr4dig, 1, 1) = "7" Then wtexto1 = wtexto1 & n07 & mil
If Mid(vlr4dig, 1, 1) = "8" Then wtexto1 = wtexto1 & n08 & mil
If Mid(vlr4dig, 1, 1) = "9" Then wtexto1 = wtexto1 & n09 & mil
If Mid(vlr4dig, 2, 3) = "000" Then
   wtexto1 = wtexto1 & moedas
Else
   wtexto1 = wtexto1 & ", "
   vlr3dig = Mid(vlr4dig, 2, 3)
   GoSub rot3dig
End If
Return

' Rotina para cinco digitos inteiros de extenso
rot5dig:
If Mid(vlr5dig, 1, 1) = "0" Then
   vlr4dig = Mid(vlr5dig, 2, 4)
   GoSub rot4dig
   Return
End If
If Val(Mid(vlr5dig, 1, 2)) < 20 Then
   If Mid(vlr5dig, 1, 2) = "10" Then wtexto1 = wtexto1 & n10 & mil
   If Mid(vlr5dig, 1, 2) = "11" Then wtexto1 = wtexto1 & n11 & mil
   If Mid(vlr5dig, 1, 2) = "12" Then wtexto1 = wtexto1 & n12 & mil
   If Mid(vlr5dig, 1, 2) = "13" Then wtexto1 = wtexto1 & n13 & mil
   If Mid(vlr5dig, 1, 2) = "14" Then wtexto1 = wtexto1 & n14 & mil
   If Mid(vlr5dig, 1, 2) = "15" Then wtexto1 = wtexto1 & n15 & mil
   If Mid(vlr5dig, 1, 2) = "16" Then wtexto1 = wtexto1 & n16 & mil
   If Mid(vlr5dig, 1, 2) = "17" Then wtexto1 = wtexto1 & n17 & mil
   If Mid(vlr5dig, 1, 2) = "18" Then wtexto1 = wtexto1 & n18 & mil
   If Mid(vlr5dig, 1, 2) = "19" Then wtexto1 = wtexto1 & n19 & mil
   If Mid(vlr5dig, 3, 3) = "000" Then
      wtexto1 = wtexto1 & moedas
   Else
      wtexto1 = wtexto1 & ", "
      vlr3dig = Mid(vlr5dig, 3, 3)
      GoSub rot3dig
   End If
   Return
Else
   If Mid(vlr5dig, 1, 1) = "1" Then wtexto1 = wtexto1 & n10
   If Mid(vlr5dig, 1, 1) = "2" Then wtexto1 = wtexto1 & n20
   If Mid(vlr5dig, 1, 1) = "3" Then wtexto1 = wtexto1 & n30
   If Mid(vlr5dig, 1, 1) = "4" Then wtexto1 = wtexto1 & n40
   If Mid(vlr5dig, 1, 1) = "5" Then wtexto1 = wtexto1 & n50
   If Mid(vlr5dig, 1, 1) = "6" Then wtexto1 = wtexto1 & n60
   If Mid(vlr5dig, 1, 1) = "7" Then wtexto1 = wtexto1 & n70
   If Mid(vlr5dig, 1, 1) = "8" Then wtexto1 = wtexto1 & n80
   If Mid(vlr5dig, 1, 1) = "9" Then wtexto1 = wtexto1 & n90
End If
If Mid(vlr5dig, 2, 4) = "0000" Then
   wtexto1 = wtexto1 & mil & moedas
   Return
End If
If Mid(vlr5dig, 2, 1) = "0" Then
   wtexto1 = wtexto1 & mil & ", "
   vlr3dig = Mid(vlr5dig, 3, 3)
   GoSub rot3dig
   Return
End If
wtexto1 = wtexto1 & " E "
vlr4dig = Mid(vlr5dig, 2, 4)
GoSub rot4dig
Return

' Rotina para Seis digitos inteiros de extenso
rot6dig:
If Mid(vlr6dig, 1, 1) = "0" Then
   vlr5dig = Mid(vlr6dig, 2, 5)
   GoSub rot5dig
   Return
End If
If Mid(vlr6dig, 1, 3) = "100" Then
   If Mid(vlr6dig, 4, 3) = "000" Then
      wtexto1 = wtexto1 & n100 & mil & moedas
   Else
      wtexto1 = wtexto1 & n100 & mil & ", "
      vlr3dig = Mid(vlr6dig, 4, 3)
      GoSub rot3dig
   End If
   Return
End If
If Mid(vlr6dig, 1, 1) = "1" Then wtexto1 = wtexto1 & n100a
If Mid(vlr6dig, 1, 1) = "2" Then wtexto1 = wtexto1 & n200
If Mid(vlr6dig, 1, 1) = "3" Then wtexto1 = wtexto1 & n300
If Mid(vlr6dig, 1, 1) = "4" Then wtexto1 = wtexto1 & n400
If Mid(vlr6dig, 1, 1) = "5" Then wtexto1 = wtexto1 & n500
If Mid(vlr6dig, 1, 1) = "6" Then wtexto1 = wtexto1 & n600
If Mid(vlr6dig, 1, 1) = "7" Then wtexto1 = wtexto1 & n700
If Mid(vlr6dig, 1, 1) = "8" Then wtexto1 = wtexto1 & n800
If Mid(vlr6dig, 1, 1) = "9" Then wtexto1 = wtexto1 & n900
If Mid(vlr6dig, 2, 5) = "00000" Then
   wtexto1 = wtexto1 & mil & moedas
   Return
End If
If Mid(vlr6dig, 2, 2) = "00" Then
   wtexto1 = wtexto1 & mil & ", "
   vlr3dig = Mid(vlr6dig, 4, 3)
   GoSub rot3dig
   Return
End If
wtexto1 = wtexto1 & " E "
vlr5dig = Mid(vlr6dig, 2, 5)
GoSub rot5dig
Return

' Rotina para Sete digitos inteiros de extenso
rot7dig:
If Mid(vlr7dig, 1, 1) = "0" Then
   vlr6dig = Mid(vlr7dig, 2, 6)
   GoSub rot6dig
   Return
End If
If Mid(vlr7dig, 1, 1) = "1" Then wtexto1 = wtexto1 & n01 & milhao
If Mid(vlr7dig, 1, 1) = "2" Then wtexto1 = wtexto1 & n02 & milhoes
If Mid(vlr7dig, 1, 1) = "3" Then wtexto1 = wtexto1 & n03 & milhoes
If Mid(vlr7dig, 1, 1) = "4" Then wtexto1 = wtexto1 & n04 & milhoes
If Mid(vlr7dig, 1, 1) = "5" Then wtexto1 = wtexto1 & n05 & milhoes
If Mid(vlr7dig, 1, 1) = "6" Then wtexto1 = wtexto1 & n06 & milhoes
If Mid(vlr7dig, 1, 1) = "7" Then wtexto1 = wtexto1 & n07 & milhoes
If Mid(vlr7dig, 1, 1) = "8" Then wtexto1 = wtexto1 & n08 & milhoes
If Mid(vlr7dig, 1, 1) = "9" Then wtexto1 = wtexto1 & n09 & milhoes
If Mid(vlr7dig, 2, 6) = "000000" Then
   wtexto1 = wtexto1 & " DE " & moedas
Else
   wtexto1 = wtexto1 & ", "
   vlr6dig = Mid(vlr7dig, 2, 6)
   GoSub rot6dig
End If
Return

' Rotina para Oito digitos inteiros de extenso
rot8dig:
If Val(Mid(vlr8dig, 1, 2)) < 20 Then
   If Mid(vlr8dig, 1, 2) = "10" Then wtexto1 = wtexto1 & n10 & milhoes
   If Mid(vlr8dig, 1, 2) = "11" Then wtexto1 = wtexto1 & n11 & milhoes
   If Mid(vlr8dig, 1, 2) = "12" Then wtexto1 = wtexto1 & n12 & milhoes
   If Mid(vlr8dig, 1, 2) = "13" Then wtexto1 = wtexto1 & n13 & milhoes
   If Mid(vlr8dig, 1, 2) = "14" Then wtexto1 = wtexto1 & n14 & milhoes
   If Mid(vlr8dig, 1, 2) = "15" Then wtexto1 = wtexto1 & n15 & milhoes
   If Mid(vlr8dig, 1, 2) = "16" Then wtexto1 = wtexto1 & n16 & milhoes
   If Mid(vlr8dig, 1, 2) = "17" Then wtexto1 = wtexto1 & n17 & milhoes
   If Mid(vlr8dig, 1, 2) = "18" Then wtexto1 = wtexto1 & n18 & milhoes
   If Mid(vlr8dig, 1, 2) = "19" Then wtexto1 = wtexto1 & n19 & milhoes
   If Mid(vlr8dig, 2, 7) = "0000000" Then
      wtexto1 = wtexto1 & " DE " & moedas
   Else
      wtexto1 = wtexto1 & ", "
      vlr6dig = Mid(vlr8dig, 3, 6)
      GoSub rot6dig
   End If
   Return
Else
   If Mid(vlr8dig, 1, 1) = "1" Then wtexto1 = wtexto1 & n10
   If Mid(vlr8dig, 1, 1) = "2" Then wtexto1 = wtexto1 & n20
   If Mid(vlr8dig, 1, 1) = "3" Then wtexto1 = wtexto1 & n30
   If Mid(vlr8dig, 1, 1) = "4" Then wtexto1 = wtexto1 & n40
   If Mid(vlr8dig, 1, 1) = "5" Then wtexto1 = wtexto1 & n50
   If Mid(vlr8dig, 1, 1) = "6" Then wtexto1 = wtexto1 & n60
   If Mid(vlr8dig, 1, 1) = "7" Then wtexto1 = wtexto1 & n70
   If Mid(vlr8dig, 1, 1) = "8" Then wtexto1 = wtexto1 & n80
   If Mid(vlr8dig, 1, 1) = "9" Then wtexto1 = wtexto1 & n90
End If
If Mid(vlr8dig, 2, 7) = "0000000" Then
   wtexto1 = wtexto1 & milhoes & " DE " & moedas
Else
   If Mid(vlr8dig, 2, 1) = "0" Then
      wtexto1 = wtexto1 & milhoes & ", "
      vlr6dig = Mid(vlr8dig, 3, 6)
      GoSub rot6dig
      Return
   Else
      wtexto1 = wtexto1 & " E "
      vlr7dig = Mid(vlr8dig, 2, 7)
   End If
   GoSub rot7dig
End If
Return

' Rotina para gerar o extenso do decimal
rot1digdec:
If dec1dig = "1" Then
   If Val(dec1dig) = 1 Then
      wtexto2 = wtexto2 & n01 & centavo
   Else
     wtexto2 = wtexto2 & n01 & centavos
   End If
ElseIf dec1dig = "2" Then
   wtexto2 = wtexto2 & n02 & centavos
ElseIf dec1dig = "3" Then
   wtexto2 = wtexto2 & n03 & centavos
ElseIf dec1dig = "4" Then
   wtexto2 = wtexto2 & n04 & centavos
ElseIf dec1dig = "5" Then
   wtexto2 = wtexto2 & n05 & centavos
ElseIf dec1dig = "6" Then
   wtexto2 = wtexto2 & n06 & centavos
ElseIf dec1dig = "7" Then
   wtexto2 = wtexto2 & n07 & centavos
ElseIf dec1dig = "8" Then
   wtexto2 = wtexto2 & n08 & centavos
ElseIf dec1dig = "9" Then
   wtexto2 = wtexto2 & n09 & centavos
End If
Return

' Rotina para dois digitos inteiros de extenso
rot2digdec:
If Mid(dec2dig, 1, 1) = "0" Then
   dec1dig = Mid(dec2dig, 2, 1)
   GoSub rot1digdec
   wtexto2 = " E " & wtexto2
   Return
End If
If Val(dec2dig) < 20 Then
   If dec2dig = "10" Then wtexto2 = wtexto2 & n10 & centavos
   If dec2dig = "11" Then wtexto2 = wtexto2 & n11 & centavos
   If dec2dig = "12" Then wtexto2 = wtexto2 & n12 & centavos
   If dec2dig = "13" Then wtexto2 = wtexto2 & n13 & centavos
   If dec2dig = "14" Then wtexto2 = wtexto2 & n14 & centavos
   If dec2dig = "15" Then wtexto2 = wtexto2 & n15 & centavos
   If dec2dig = "16" Then wtexto2 = wtexto2 & n16 & centavos
   If dec2dig = "17" Then wtexto2 = wtexto2 & n17 & centavos
   If dec2dig = "18" Then wtexto2 = wtexto2 & n18 & centavos
   If dec2dig = "19" Then wtexto2 = wtexto2 & n19 & centavos
   If Len(Trim(wtexto2)) > 0 Then wtexto2 = " E " & wtexto2
   Return
Else
   If Mid(dec2dig, 1, 1) = "2" Then wtexto2 = wtexto2 & n20
   If Mid(dec2dig, 1, 1) = "3" Then wtexto2 = wtexto2 & n30
   If Mid(dec2dig, 1, 1) = "4" Then wtexto2 = wtexto2 & n40
   If Mid(dec2dig, 1, 1) = "5" Then wtexto2 = wtexto2 & n50
   If Mid(dec2dig, 1, 1) = "6" Then wtexto2 = wtexto2 & n60
   If Mid(dec2dig, 1, 1) = "7" Then wtexto2 = wtexto2 & n70
   If Mid(dec2dig, 1, 1) = "8" Then wtexto2 = wtexto2 & n80
   If Mid(dec2dig, 1, 1) = "9" Then wtexto2 = wtexto2 & n90
End If
If Mid(dec2dig, 2, 1) = "0" Then
   wtexto2 = wtexto2 & centavos
Else
   wtexto2 = wtexto2 & " E "
   dec1dig = Mid(dec2dig, 2, 1)
   GoSub rot1digdec
End If
If Len(Trim(wtexto2)) > 0 Then
   wtexto2 = " E " & wtexto2
End If
Return
End Function

'Colocar R$ nos valores convertidos

Public Function convQtd(vlr As String, comZeros As Boolean) As String
   Dim i As Integer
   Dim tamStr As Integer
   Dim countPoint As Integer
   
   'Testar se não é tudo branco, se for nao fazer processamento
   convQtd = vlr
   vlr = Trim(vlr)
   
   If vlr = "" Then
      'Antes de sair da função, testar se é necessário colocar zeros ou não
      If comZeros Then _
         convQtd = "0,000"
      Exit Function
   End If
   
   tamStr = Len(vlr) 'Tamanho do valor passado
   i = 1
   'Enquanto for "0", preencher com " " (espaços em branco) (Fazer até tamanho - 2 para não _
    limpar os decimais, que podem ser 0
   While Mid(vlr, i, 1) = "0" And i < tamStr - 2
      Mid(vlr, i, 1) = " "
      i = i + 1
   Wend

   'Iniciar o resultado com os decimais e a vírgula ",xx"
   convQtd = "," & Right(vlr, 3)
   
   i = Len(vlr) - 3     'Tamanho total - 2
   countPoint = 0
   While i > 0
      convQtd = Mid(vlr, i, 1) & convQtd  'Pegar cada número, de trás pra frente, e colocar _
                                           na variável de saída, na frente dos decimais
      countPoint = countPoint + 1
      If i > 1 Then
         'Se o próximo valor a ser lido for um número e se o contador for 3, então colocar _
          o ponto dos decimais
         If Mid(vlr, i - 1, 1) <> " " And countPoint = 3 Then
            convQtd = "." & convQtd
            countPoint = 0
         End If
      End If
      i = i - 1   'Decremento do while
   Wend
   convQtd = Trim(convQtd)
End Function

Public Function CentraStr(Campo As String, Tamanho As Integer) As String
   Dim i As Integer
   Dim tamStr As Integer
   
  'Testar se não é tudo branco, se for nao fazer processamento
   If Len(Trim(Campo)) = 0 Then
      CentraStr = ""
      Exit Function
   End If
   CentraStr = Space(Tamanho - (Len(Trim(Campo)) / 2)) & Trim(Campo)
End Function
Public Function DireitaStr(Campo As String, Tamanho As Integer) As String
   Dim i As Integer
   Dim tamStr As Integer
   
  'Testar se não é tudo branco, se for nao fazer processamento
   If Len(Trim(Campo)) = 0 Then
      DireitaStr = ""
      Exit Function
   End If
   DireitaStr = Space(Tamanho - Len(Trim(Campo))) & Trim(Campo)
End Function

Function CodigoBarra(ByVal linha As String)
  Dim A, BB As Integer
  Dim TesIgual As String * 2
  Dim Barra As String
  
  Barra = ""
  A = 1
  
  While A < 44
    TesIgual = Mid(linha, A, 2)
    If TesIgual = "00" Then
        Barra = Barra & "nnWWn"
    ElseIf TesIgual = "01" Then
        Barra = Barra & "NnwwN"
    ElseIf TesIgual = "02" Then
        Barra = Barra & "nNwwN"
    ElseIf TesIgual = "03" Then
        Barra = Barra & "NNwwn"
    ElseIf TesIgual = "04" Then
        Barra = Barra & "nnWwN"
    ElseIf TesIgual = "05" Then
        Barra = Barra & "NnWwn"
    ElseIf TesIgual = "06" Then
        Barra = Barra & "nNWwn"
    ElseIf TesIgual = "07" Then
        Barra = Barra & "nnwWN"
    ElseIf TesIgual = "08" Then
        Barra = Barra & "NnwWn"
    ElseIf TesIgual = "09" Then
        Barra = Barra & "nNwWn"
    ElseIf TesIgual = "10" Then
        Barra = Barra & "wnNNw"
    ElseIf TesIgual = "11" Then
        Barra = Barra & "WnnnW"
    ElseIf TesIgual = "12" Then
        Barra = Barra & "wNnnW"
    ElseIf TesIgual = "13" Then
        Barra = Barra & "WNnnw"
    ElseIf TesIgual = "14" Then
        Barra = Barra & "wnNnW"
    ElseIf TesIgual = "15" Then
        Barra = Barra & "WnNnw"
    ElseIf TesIgual = "16" Then
        Barra = Barra & "wNNnw"
    ElseIf TesIgual = "17" Then
        Barra = Barra & "wnnNW"
    ElseIf TesIgual = "18" Then
        Barra = Barra & "WnnNw"
    ElseIf TesIgual = "19" Then
        Barra = Barra & "wNnNw"
    ElseIf TesIgual = "20" Then
        Barra = Barra & "nwNNw"
    ElseIf TesIgual = "21" Then
        Barra = Barra & "NwnnW"
    ElseIf TesIgual = "22" Then
        Barra = Barra & "nWnnW"
    ElseIf TesIgual = "23" Then
        Barra = Barra & "NWnnw"
    ElseIf TesIgual = "24" Then
        Barra = Barra & "nwNnW"
    ElseIf TesIgual = "25" Then
        Barra = Barra & "NwNnw"
    ElseIf TesIgual = "26" Then
        Barra = Barra & "nWNnw"
    ElseIf TesIgual = "27" Then
        Barra = Barra & "nwnNW"
    ElseIf TesIgual = "28" Then
        Barra = Barra & "NwnNw"
    ElseIf TesIgual = "29" Then
        Barra = Barra & "nWnNw"
    ElseIf TesIgual = "30" Then
        Barra = Barra & "wwNNn"
    ElseIf TesIgual = "31" Then
        Barra = Barra & "WwnnN"
    ElseIf TesIgual = "32" Then
        Barra = Barra & "wWnnN"
    ElseIf TesIgual = "33" Then
        Barra = Barra & "WWnnn"
    ElseIf TesIgual = "34" Then
        Barra = Barra & "wwNnN"
    ElseIf TesIgual = "35" Then
        Barra = Barra & "WwNnn"
    ElseIf TesIgual = "36" Then
        Barra = Barra & "wWNnn"
    ElseIf TesIgual = "37" Then
        Barra = Barra & "wwnNN"
    ElseIf TesIgual = "38" Then
        Barra = Barra & "WwnNn"
    ElseIf TesIgual = "39" Then
        Barra = Barra & "wWnNn"
    ElseIf TesIgual = "40" Then
        Barra = Barra & "nnWNw"
    ElseIf TesIgual = "41" Then
        Barra = Barra & "NnwnW"
    ElseIf TesIgual = "42" Then
        Barra = Barra & "nNwnW"
    ElseIf TesIgual = "43" Then
        Barra = Barra & "NNwnw"
    ElseIf TesIgual = "44" Then
        Barra = Barra & "nnWnW"
    ElseIf TesIgual = "45" Then
        Barra = Barra & "NnWnw"
    ElseIf TesIgual = "46" Then
        Barra = Barra & "nNWnw"
    ElseIf TesIgual = "47" Then
        Barra = Barra & "nnwNW"
    ElseIf TesIgual = "48" Then
        Barra = Barra & "NnwNw"
    ElseIf TesIgual = "49" Then
        Barra = Barra & "nNwNw"
    ElseIf TesIgual = "50" Then
        Barra = Barra & "wnWNn"
    ElseIf TesIgual = "51" Then
        Barra = Barra & "WnwnN"
    ElseIf TesIgual = "52" Then
        Barra = Barra & "wNwnN"
    ElseIf TesIgual = "53" Then
        Barra = Barra & "WNwnn"
    ElseIf TesIgual = "54" Then
        Barra = Barra & "wnWnN"
    ElseIf TesIgual = "55" Then
        Barra = Barra & "WnWnn"
    ElseIf TesIgual = "56" Then
        Barra = Barra & "wNWnn"
    ElseIf TesIgual = "57" Then
        Barra = Barra & "wnwNN"
    ElseIf TesIgual = "58" Then
        Barra = Barra & "WnwNn"
    ElseIf TesIgual = "59" Then
        Barra = Barra & "wNwNn"
    ElseIf TesIgual = "60" Then
        Barra = Barra & "nwWNn"
    ElseIf TesIgual = "61" Then
        Barra = Barra & "NwwnN"
    ElseIf TesIgual = "62" Then
        Barra = Barra & "nWwnN"
    ElseIf TesIgual = "63" Then
        Barra = Barra & "NWwnn"
    ElseIf TesIgual = "64" Then
        Barra = Barra & "nwWnN"
    ElseIf TesIgual = "65" Then
        Barra = Barra & "NwWnn"
    ElseIf TesIgual = "66" Then
        Barra = Barra & "nWWnn"
    ElseIf TesIgual = "67" Then
        Barra = Barra & "nwwNN"
    ElseIf TesIgual = "68" Then
        Barra = Barra & "NwwNn"
    ElseIf TesIgual = "69" Then
        Barra = Barra & "nWwNn"
    ElseIf TesIgual = "70" Then
        Barra = Barra & "nnNWw"
    ElseIf TesIgual = "71" Then
        Barra = Barra & "NnnwW"
    ElseIf TesIgual = "72" Then
        Barra = Barra & "nNnwW"
    ElseIf TesIgual = "73" Then
        Barra = Barra & "NNnww"
    ElseIf TesIgual = "74" Then
        Barra = Barra & "nnNwW"
    ElseIf TesIgual = "75" Then
        Barra = Barra & "NnNww"
    ElseIf TesIgual = "76" Then
        Barra = Barra & "nNNww"
    ElseIf TesIgual = "77" Then
        Barra = Barra & "nnnWW"
    ElseIf TesIgual = "78" Then
        Barra = Barra & "NnnWw"
    ElseIf TesIgual = "79" Then
        Barra = Barra & "nNnWw"
    ElseIf TesIgual = "80" Then
        Barra = Barra & "wnNWn"
    ElseIf TesIgual = "81" Then
        Barra = Barra & "WnnwN"
    ElseIf TesIgual = "82" Then
        Barra = Barra & "wNnwN"
    ElseIf TesIgual = "83" Then
        Barra = Barra & "WNnwn"
    ElseIf TesIgual = "84" Then
        Barra = Barra & "wnNwN"
    ElseIf TesIgual = "85" Then
        Barra = Barra & "WnNwn"
    ElseIf TesIgual = "86" Then
        Barra = Barra & "wNNwn"
    ElseIf TesIgual = "87" Then
        Barra = Barra & "wnnWN"
    ElseIf TesIgual = "88" Then
        Barra = Barra & "WnnWn"
    ElseIf TesIgual = "89" Then
        Barra = Barra & "wNnWn"
    ElseIf TesIgual = "90" Then
        Barra = Barra & "nwNWn"
    ElseIf TesIgual = "91" Then
        Barra = Barra & "NwnwN"
    ElseIf TesIgual = "92" Then
        Barra = Barra & "nWnwN"
    ElseIf TesIgual = "93" Then
        Barra = Barra & "NWnwn"
    ElseIf TesIgual = "94" Then
        Barra = Barra & "nwNwN"
    ElseIf TesIgual = "95" Then
        Barra = Barra & "NwNwn"
    ElseIf TesIgual = "96" Then
        Barra = Barra & "nWNwn"
    ElseIf TesIgual = "97" Then
        Barra = Barra & "nwnWN"
    ElseIf TesIgual = "98" Then
        Barra = Barra & "NwnWn"
    ElseIf TesIgual = "99" Then
        Barra = Barra & "nWnWn"
    End If
    A = A + 2
  Wend
  BB = Len(Barra)
  CodigoBarra = Barra
End Function

Function moduloCEP(ByVal pNum As String)
Dim IND As Integer
Dim lVlPro As Integer
Dim lVlTeste As String
Dim lCampo As String
Dim wtam As Integer
Dim wtam1 As Integer
Dim wvar As Integer
Dim wtotal As Integer
Dim i As Integer
 
        lCampo = pNum
        lVlTeste = lCampo
        wtam = Len(lVlTeste)
        wtam1 = Len(lVlTeste)
        wvar = 2
        lVlPro = 0
        wtotal = 0
        For i = 1 To wtam
            lVlPro = Val(Mid(lVlTeste, wtam1, 1))
            wtotal = wtotal + lVlPro
            If wvar = 2 Then
               wvar = 1
            Else
               wvar = 2
            End If
            wtam1 = wtam1 - 1
        Next i
        If wtotal > 9 Then
           wtotal = wtotal Mod 10
           If wtotal < 10 And wtotal <> 0 Then
              wtotal = 10 - wtotal
           Else
              wtotal = 0
           End If
        Else
           wtotal = 10 - wtotal
        End If
        moduloCEP = wtotal
End Function

Function modulo10(ByVal pNum As String) ' Digito verificador
Dim IND As Integer
Dim lVlPro As Integer
Dim lVlTeste As String
Dim lCampo As String
Dim wtam As Integer
Dim wtam1 As Integer
Dim wvar As Integer
Dim wtotal As Integer
Dim i As Integer

        lCampo = pNum
        lVlTeste = lCampo
        wtam = Len(lVlTeste)
        wtam1 = Len(lVlTeste)
        wvar = 2
        lVlPro = 0
        wtotal = 0
        For i = 1 To wtam
            lVlPro = Val(Mid(lVlTeste, wtam1, 1)) * wvar
            If lVlPro > 9 Then
               lVlPro = Val(Mid(lVlPro, 1, 1)) + Val(Mid(lVlPro, 2, 1))
            End If
            wtotal = wtotal + lVlPro
            If wvar = 2 Then
               wvar = 1
            Else
               wvar = 2
            End If
            wtam1 = wtam1 - 1
        Next i
        If wtotal > 9 Then
           wtotal = wtotal Mod 10
           If wtotal < 10 And wtotal <> 0 Then
              wtotal = 10 - wtotal
           Else
              wtotal = 0
           End If
        Else
           wtotal = 10 - wtotal
        End If
        modulo10 = wtotal
End Function

Public Function Modulo11(ByVal pNum As String) As Integer
   Dim IND As Integer
   Dim lVlPro As Integer
   Dim lVlTeste As String
   Dim lCampo As String
   Dim valfator As Integer
   Dim lAjuste As Integer
   
   lVlPro = 0
   IND = Len(pNum) + 1
   If IND = 44 Then
      lAjuste = 1
   ElseIf IND = 11 Then
      lAjuste = 0
   Else
      lAjuste = 1
   End If
   valfator = 1
   lCampo = pNum
   lVlTeste = pNum
   While IND > 1
      valfator = valfator + 1
      IND = IND - 1
      lVlPro = lVlPro + Mid$(lVlTeste, IND, 1) * valfator
      If valfator = 9 Then
         valfator = 1
      End If
   Wend
   lVlPro = lVlPro Mod 11
   lVlPro = 11 - lVlPro
   If lVlPro = 10 Or lVlPro = 11 Or lVlPro = 0 Then
      lVlPro = lAjuste
   End If
   Modulo11 = lVlPro
End Function
 

Public Function MesAno(dt As String) As String
Dim mes(12) As String * 20
   mes(0) = ""
   mes(1) = "Janeiro"
   mes(2) = "Fevereiro"
   mes(3) = "Março"
   mes(4) = "Abril"
   mes(5) = "Maio"
   mes(6) = "Junho"
   mes(7) = "Julho"
   mes(8) = "Agosto"
   mes(9) = "Setembro"
   mes(10) = "Outubro"
   mes(11) = "Novembro"
   mes(12) = "Dezembro"
   MesAno = Trim(mes(CInt(Mid(dt, 4, 2)))) & " / " & Right(dt, 4)
End Function
Public Function Code128c(wstring As String)
Dim sBarCodeTable(100) As String
    sBarCodeTable(0) = " "
    sBarCodeTable(1) = "!"
    sBarCodeTable(2) = Chr(34)
    sBarCodeTable(3) = "#"
    sBarCodeTable(4) = "$"
    sBarCodeTable(5) = "%"
    sBarCodeTable(6) = "&"
    sBarCodeTable(7) = "'"
    sBarCodeTable(8) = "\("
    sBarCodeTable(9) = "\)"
    sBarCodeTable(10) = "*"
    sBarCodeTable(11) = "+"
    sBarCodeTable(12) = ","
    sBarCodeTable(13) = "-"
    sBarCodeTable(14) = "."
    sBarCodeTable(15) = "/"
    sBarCodeTable(16) = "0"
    sBarCodeTable(17) = "1"
    sBarCodeTable(18) = "2"
    sBarCodeTable(19) = "3"
    sBarCodeTable(20) = "4"
    sBarCodeTable(21) = "5"
    sBarCodeTable(22) = "6"
    sBarCodeTable(23) = "7"
    sBarCodeTable(24) = "8"
    sBarCodeTable(25) = "9"
    sBarCodeTable(26) = ":"
    sBarCodeTable(27) = ";"
    sBarCodeTable(28) = "<"
    sBarCodeTable(29) = "="
    sBarCodeTable(30) = ">"
    sBarCodeTable(31) = "?"
    sBarCodeTable(32) = "@"
    sBarCodeTable(33) = "A"
    sBarCodeTable(34) = "B"
    sBarCodeTable(35) = "C"
    sBarCodeTable(36) = "D"
    sBarCodeTable(37) = "E"
    sBarCodeTable(38) = "F"
    sBarCodeTable(39) = "G"
    sBarCodeTable(40) = "H"
    sBarCodeTable(41) = "I"
    sBarCodeTable(42) = "J"
    sBarCodeTable(43) = "K"
    sBarCodeTable(44) = "L"
    sBarCodeTable(45) = "M"
    sBarCodeTable(46) = "N"
    sBarCodeTable(47) = "O"
    sBarCodeTable(48) = "P"
    sBarCodeTable(49) = "Q"
    sBarCodeTable(50) = "R"
    sBarCodeTable(51) = "S"
    sBarCodeTable(52) = "T"
    sBarCodeTable(53) = "U"
    sBarCodeTable(54) = "V"
    sBarCodeTable(55) = "W"
    sBarCodeTable(56) = "X"
    sBarCodeTable(57) = "Y"
    sBarCodeTable(58) = "Z"
    sBarCodeTable(59) = "["
    sBarCodeTable(60) = "\\"
    sBarCodeTable(61) = "]"
    sBarCodeTable(62) = "^"
    sBarCodeTable(63) = "_"
    sBarCodeTable(64) = "`"
    sBarCodeTable(65) = "a"
    sBarCodeTable(66) = "b"
    sBarCodeTable(67) = "c"
    sBarCodeTable(68) = "d"
    sBarCodeTable(69) = "e"
    sBarCodeTable(70) = "f"
    sBarCodeTable(71) = "g"
    sBarCodeTable(72) = "h"
    sBarCodeTable(73) = "i"
    sBarCodeTable(74) = "j"
    sBarCodeTable(75) = "k"
    sBarCodeTable(76) = "l"
    sBarCodeTable(77) = "m"
    sBarCodeTable(78) = "n"
    sBarCodeTable(79) = "o"
    sBarCodeTable(80) = "p"
    sBarCodeTable(81) = "q"
    sBarCodeTable(82) = "r"
    sBarCodeTable(83) = "s"
    sBarCodeTable(84) = "t"
    sBarCodeTable(85) = "u"
    sBarCodeTable(86) = "v"
    sBarCodeTable(87) = "w"
    sBarCodeTable(88) = "x"
    sBarCodeTable(89) = "y"
    sBarCodeTable(90) = "z"
    sBarCodeTable(91) = "{"
    sBarCodeTable(92) = "|"
    sBarCodeTable(93) = "}"
    sBarCodeTable(94) = "~"
    sBarCodeTable(95) = Chr(127)
    sBarCodeTable(96) = Chr(128)
    sBarCodeTable(97) = Chr(129)
    sBarCodeTable(98) = Chr(130)
    sBarCodeTable(99) = Chr(131)
    Code128c = ""
    Dim i As Integer
    For i = 1 To Len(wstring) Step 2
      Code128c = Code128c & sBarCodeTable(Mid(wstring, i, 2))
    Next
End Function
Public Function criptografa(wsen As String) As String
       Dim wNova As String
       Dim i As Integer
       wNova = ""
       wsen = Trim(wsen)
       For i = 1 To Len(wsen)
           wNova = wNova + Chr(Asc(Mid(wsen, i, 1)) + 5)
       Next
       criptografa = wNova
End Function
Public Function Descriptografa(wsen As String) As String
       Dim wNova As String
       Dim i As Integer
       wNova = ""
       wsen = Trim(wsen)
       For i = 1 To Len(wsen)
           wNova = wNova + Chr(Asc(Mid(wsen, i, 1)) - 5)
       Next
       Descriptografa = wNova
End Function
