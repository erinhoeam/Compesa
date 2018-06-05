VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdem_Fisc_Inat_Compesa_A5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap Consultoria
'                 Analista de Sistemas
'Data: 26 de Maio de 2009
'Contatos - Email: erinhoeam@hotmail.com
'                  aziel@sistemap.com.br
'           Fone: (81) 9832.6084
'Aplicação: Ordem de Fiscalização Inativo Compesa A5
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double
Dim wQtdFaturas As Long
Dim wIndiceFaturas As Long

Const AppName = "Compesa", AppDesc = "Ordem de Fiscalização Inativo Compesa A5"

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
        TOTAL_CATEGORIA As String
        PERFIL_CLIENTE As String
        GRUPO_DE_COBRANCA As String
        SEQUENCIAL As String
        NR_ORDEM As String
        GRUPO As String
        RES_COM_IND_PUB As String
        TOTAL_ECONOMIA As String
        DATA_EMISSAO As String
        TIPO_CONSUMIDOR As String
        ULTIMA_ALTERACAO As String
        SITUACAO_AGUA As String
        CONSUMO_MEDIO As String
        SITUACAO_ESGOTO As String
        CONSUMO_FIXO As String
        DATA_CORTE As String
        DATA_SUPRESSAO As String
        ORIGEM As String
        OCORRENCIA As String
End Type
'Variáveis gerais para os tipos
Private Reg(2) As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Long
        Dim posY As Long
        Dim posX As Long
        wRec = wRec + 1

        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "90 rotate 00 pontos 00 pontos translate" ' A4 LANDSCAPE
        For i = 1 To 2
            Select Case i
                Case 1
                    posX = 100
                    posY = 305
                Case 2
                    posX = 1860
                    posY = 305
            End Select
            
            If Len(Trim(Reg(i).NOME)) > 0 Then totRegs = totRegs + 1
   
            'Inscrição, Nome
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).INSCRICAO & ") show"
            Print #outFN, posX + 610 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg(i).MATRICULA & ") show"
            posY = posY + 80
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            'Endereço, Grupo, Numero de Ordem
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).ENDERECO & "  " & Reg(i).BAIRRO & " - " & Reg(i).MUNICIPIO & "/" & Reg(i).UF & ") show"
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
            Print #outFN, posX + 1380 & " pontos -" & posY & " pontos moveto (" & Reg(i).NR_ORDEM & ") show"
            posY = posY + 85
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 90 & " pontos -" & posY & " pontos moveto (" & Reg(i).RES & ") show"
            Print #outFN, posX + 180 & " pontos -" & posY & " pontos moveto (" & Reg(i).COM & ") show"
            Print #outFN, posX + 270 & " pontos -" & posY & " pontos moveto (" & Reg(i).IND & ") show"
            Print #outFN, posX + 340 & " pontos -" & posY & " pontos moveto (" & Reg(i).PUB & ") show"
            Print #outFN, posX + 445 & " pontos -" & posY & " pontos moveto (" & Reg(i).TOTAL_CATEGORIA & ") show"
            Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_EMISSAO & ") show"
            Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg(i).TIPO_CONSUMIDOR & ") show"
            Print #outFN, posX + 1070 & " pontos -" & posY & " pontos moveto (" & Reg(i).ULTIMA_ALTERACAO & ") show"
            Print #outFN, posX + 1370 & " pontos -" & posY & " pontos moveto (" & Reg(i).GRUPO & ") show"
            Print #outFN, posX + 1480 & " pontos -" & posY & " pontos moveto (" & Reg(i).SEQUENCIAL & ") show"
            posY = posY + 85
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).SITUACAO_AGUA & ") show"
            Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg(i).CONSUMO_MEDIO & ") show"
            Print #outFN, posX + 830 & " pontos -" & posY & " pontos moveto (" & Reg(i).SITUACAO_ESGOTO & ") show"
            posY = posY + 90
            Print #outFN, posX + 35 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_CORTE & ") show"
            Print #outFN, posX + 550 & " pontos -" & posY & " pontos moveto (" & Reg(i).DATA_SUPRESSAO & ") show"
            Print #outFN, posX + 790 & " pontos -" & posY & " pontos moveto (" & Reg(i).ORIGEM & ") show"
            Print #outFN, posX + 980 & " pontos -" & posY & " pontos moveto (" & Reg(i).OCORRENCIA & ") show"
            posY = posY + 870
            Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 330 & " pontos -" & posY & " pontos moveto (" & Reg(i).NOME & ") show"
            posY = posY + 950
            Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
            Print #outFN, posX + 50 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
            Print #outFN, posX + 200 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Next
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "
        
        'If wRec = 20 Then End
End Sub
Private Sub Processar()
        Screen.MousePointer = vbHourglass
        Dim linha As String, i As Integer
        Dim wIndice As Integer
        wIndice = 0
        totRegs = 0

        'Gravar Cabeçalho
        GravaHeader

        wArq = frmPrincipal.LstSpool.List(frmPrincipal.LstSpool.ListIndex)
        'frmPrincipal.ProgressBarProcess.Max = LOF(inFN)
        'frmPrincipal.ProgressBarProcess.Value = 0.0001

        Do While Not EOF(inFN)
           DoEvents
           Line Input #inFN, linha
           wIndice = wIndice + 1

           linha = Replace(linha, "(", " ")
           linha = Replace(linha, ")", " ")
           linha = Replace(linha, "[", " ")
           linha = Replace(linha, "]", " ")

           Reg(wIndice).INSCRICAO = Trim(Mid(linha, 1, 20))
           Reg(wIndice).NOME = Trim(Mid(linha, 21, 50))
           Reg(wIndice).ENDERECO = Trim(Mid(linha, 21, 50))
           Reg(wIndice).MATRICULA = Trim(Mid(linha, 71, 9))
           Reg(wIndice).ENDERECO = Trim(Mid(linha, 80, 70))
           Reg(wIndice).BAIRRO = Trim(Mid(linha, 150, 30))
           Reg(wIndice).MUNICIPIO = Trim(Mid(linha, 180, 30))
           Reg(wIndice).UF = Trim(Mid(linha, 210, 2))
           Reg(wIndice).CEP = Trim(Mid(linha, 212, 9))
           Reg(wIndice).N_OS = Trim(Mid(linha, 221, 9))
           Reg(wIndice).RES = Trim(Mid(linha, 230, 3))
           Reg(wIndice).COM = Trim(Mid(linha, 233, 3))
           Reg(wIndice).IND = Trim(Mid(linha, 236, 3))
           Reg(wIndice).PUB = Trim(Mid(linha, 239, 3))
           Reg(wIndice).TOTAL_CATEGORIA = Trim(Mid(linha, 242, 3))
           Reg(wIndice).DATA_EMISSAO = Trim(Mid(linha, 245, 10))
           Reg(wIndice).PERFIL_CLIENTE = Trim(Mid(linha, 255, 20))
           Reg(wIndice).GRUPO_DE_COBRANCA = Trim(Mid(linha, 285, 2))
           Reg(wIndice).SEQUENCIAL = Trim(Mid(linha, 287, 7))
           Reg(wIndice).SITUACAO_AGUA = Trim(Mid(linha, 294, 20))
           Reg(wIndice).CONSUMO_MEDIO = Trim(Mid(linha, 314, 1))
           Reg(wIndice).SITUACAO_ESGOTO = Trim(Mid(linha, 342, 20))
           Reg(wIndice).CONSUMO_FIXO = Trim(Mid(linha, 343, 10))
           Reg(wIndice).DATA_CORTE = Trim(Mid(linha, 372, 10))
           Reg(wIndice).DATA_SUPRESSAO = Trim(Mid(linha, 382, 10))
           Reg(wIndice).ORIGEM = Trim(Mid(linha, 392, 15))
           Reg(wIndice).OCORRENCIA = Trim(Mid(linha, 407, 34))

           If wIndice = 2 Then
              MontaExtrato
              LimpaReg
              wIndice = 0
           End If

         '  If Seek(inFN) <= frmPrincipal.ProgressBarProcess.Max Then _
         '     frmPrincipal.ProgressBarProcess.Value = Seek(inFN)

        Loop
        Screen.MousePointer = vbDefault

        frmResultados.txtTotCli_STL.Text = totRegs
        frmResultados.txtTotReg_STL.Text = wRec
        frmResultados.txArquivo.Text = wArq

        frmResultados.txtAplicacao = AppDesc

        Close #outFN
        Close #inFN

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
        Open App.Path & "\Ordem_Ficalizacao_Inativo.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Ordem de Fiscalização Compesa A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        For i = 1 To 2
            With Reg(i)
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
                .TOTAL_CATEGORIA = ""
                .DATA_EMISSAO = ""
                .PERFIL_CLIENTE = ""
                .ULTIMA_ALTERACAO = ""
                .GRUPO_DE_COBRANCA = ""
                .SEQUENCIAL = ""
                .NR_ORDEM = ""
                .GRUPO = ""
                .RES_COM_IND_PUB = ""
                .TOTAL_ECONOMIA = ""
                .DATA_EMISSAO = ""
                .TIPO_CONSUMIDOR = ""
                .ULTIMA_ALTERACAO = ""
                .SITUACAO_AGUA = ""
                .CONSUMO_MEDIO = ""
                .SITUACAO_ESGOTO = ""
                .CONSUMO_FIXO = ""
                .DATA_CORTE = ""
                .DATA_SUPRESSAO = ""
                .ORIGEM = ""
                .OCORRENCIA = ""
            End With
        Next
End Sub
