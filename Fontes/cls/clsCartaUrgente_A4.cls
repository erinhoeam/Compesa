VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartaUrgente_A4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap Consultoria
'                 Analista de Sistemas
'Data: 18 de Novembro de 2009
'Contatos - Email: aziel@sistemap.com.br
'                  erinhoeam@hotmail.com
'           Fone: (81) 9832.6084
'Aplicação: Aviso de Cobranca Final do Ano Compesa A4
'****************************************************
Option Explicit
Dim wArq As String, wRec As Double

Const AppName = "COMPESA", AppDesc = "Carta Urgente A4"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        SEQUENCIAL As String
        MATRICULA As String
        NOME As String
        ENDERECO As String
        BAIRRO As String
        CIDADE As String
        UF As String
        CEP As String
        DATA As String
End Type
'Variáveis gerais para os tipos
Private Reg As TipoReg

'Totais de cartas geradas, cartas lidas e Registros
Private totRegs As Long
Private Sub MontaExtrato()
        Dim i As Integer
        Dim i2 As Integer
        Dim posY As Long
        Dim posX As Long
        
        wRec = wRec + 1
        posX = 260
        posY = 475
        
        '********************Parte Interna da Conta*************************************
        Print #outFN, "% Ini-" & Format(CStr(wRec), "00000000000")
        Print #outFN, "save EPSForm1 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        'Rota
'        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
'        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
'        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
'        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
'        posY = posY + 117
'        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
'        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
'        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
'        Print #outFN, posX + 1700 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        Print #outFN, "showpage"
        'Fim da Parte Interna
        '**********Inicio Parte Externa************************************************
        posX = 690
        posY = 530
        Print #outFN, "save EPSForm2 execform restore"
        Print #outFN, "00 rotate 00 pontos 3507 pontos translate" ' A4 PORTRAIT
        Print #outFN, "/Helvetica-Condensed-AP findfont 8 scalefont setfont"
        'Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO_EMPREITEIRA & ") show"
        'Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        posY = posY + 38
        'Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.ESCRITORIO & ") show"
        posY = posY + 38
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8.5 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 38
        Print #outFN, "/Helvetica-Condensed-AP findfont 7 scalefont setfont"
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
'        posY = posY + 38
'        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.BAIRRO & "      " & Reg.CIDADE & ") show"
'        Print #outFN, posX + 700 & " pontos -" & posY & " pontos moveto (" & Reg.UF & ") show"
        posY = posY + 38
        Print #outFN, posX & " pontos -" & posY & " pontos moveto (" & Reg.CEP & ") show"
'        Print #outFN, posX + 800 & " pontos -" & posY & " pontos moveto (" & Reg.SEQUENCIAL & ") show"
        posY = posY + 38
        posY = posY + 80
        Print #outFN, "/Helvetica-Condensed-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (Rec.: " & Format(wRec, "00000") & ") show"
        Print #outFN, posX + 600 & " pontos -" & posY & " pontos moveto (Arq.: " & wArq & ") show"
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "

        'If wRec = 5 Then End
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
           
'           Reg.SEQUENCIAL = Trim(Mid(linha, 1, 6))
'           Reg.MATRICULA = Trim(Mid(linha, 7, 10))
'           Reg.NOME = Trim(Mid(linha, 17, 40))
'           Reg.ENDERECO = Trim(Mid(linha, 57, 60))
'           Reg.BAIRRO = Trim(Mid(linha, 117, 50))
'           Reg.CIDADE = Trim(Mid(linha, 167, 11))
'           'Reg.UF = Trim(Mid(linha, 222, 2))
'           Reg.CEP = Trim(Mid(linha, 178, 9))
'           Reg.DATA = Trim(Mid(linha, 187, 40))

           
           Reg.NOME = Trim(Mid(linha, 1, 50))
           Reg.ENDERECO = Trim(Mid(linha, 51, 105))
           Reg.CEP = Trim(Mid(linha, 156, 9))

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
        Open App.Path & "\Carta_Urgente_A4_Int.ps" For Input As #200
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
        Open App.Path & "\Carta_Urgente_A4_Ext.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Aviso de Cobrança Compesa A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        Dim i As Integer
        Dim i2 As Integer

        Reg.SEQUENCIAL = ""
        Reg.MATRICULA = ""
        Reg.NOME = ""
        Reg.ENDERECO = ""
        Reg.BAIRRO = ""
        Reg.CIDADE = ""
        Reg.UF = ""
        Reg.CEP = ""
        Reg.DATA = ""
End Sub


