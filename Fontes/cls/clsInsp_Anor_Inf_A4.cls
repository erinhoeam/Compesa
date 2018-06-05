VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsp_Anor_Inf_A4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Aziel R. Pereira Jr.
'                 Analista de Sistemas
'Data: 31 de Agosto de 2011
'Contatos - Email: aziel@sistemap.com.br
'                  azieljr@hotmail.com
'           Fone: (81) 9979.1972
'Aplicação: Formulário de Anormalidade Informada A4
'*****************************************************
Option Explicit
Dim wArq As String, wRec As Double, wpv As Integer, i As Integer

Const AppName = "Compesa", AppDesc = "Formulário de Anormalidade Informada"

'Variáveis relacionadas com os Arquivos de Entrada e Saída
Private inFN As Integer, outFN As Integer
'Pastas que contém os arquivos
Private pastaProcessar As String, pastaGerar As String
'Arquivos que estão sendo processados
Private procArq As String, gerArq As String
'Nome do Job
Private jobName As String

Private Type TipoReg  'Dados do Beneficiario e Fatura
        N_OS              As String
        DATA_GERACAO      As String
        SEQ_OS            As String
        
        INSCRICAO         As String
        MATRICULA         As String
        PERFIL            As String
        ENDERECO          As String
        RES               As String
        COM               As String
        IND               As String
        PUB               As String
        ULTIMA_ALTERACAO  As String
        GRUPO             As String
        SITUACAO_AGUA     As String
        CONSUMO_MEDIO     As String
        SITUACAO_ESGOTO   As String
        CONSUMO_FIXO      As String
        ANORMALIDADE      As String
        
        NOME              As String
        CPF_CNPJ          As String
        RG                As String
        FONE              As String
End Type

Private Type TipoRegCSV  'Dados do Beneficiario e Fatura
        C01 As String
        C02 As String
        C03 As String
        C04 As String
        C05 As String
        C06 As String
        C07 As String
        C08 As String
        C09 As String
        C10 As String
        C11 As String
        C12 As String
        C13 As String
        C14 As String
        C15 As String
        C16 As String
        C17 As String
        C18 As String
        C19 As String
        C20 As String
        C21 As String
        C22 As String
        C23 As String
        C24 As String
        C25 As String
        C26 As String
        C27 As String
        C28 As String
        C29 As String
        C30 As String
        C31 As String
End Type
'Variáveis gerais para os tipos
Private Reg As TipoReg
Private RegCSV As TipoRegCSV

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
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1505 & " pontos -" & posY & " pontos moveto (" & Reg.N_OS & ") show"
        posY = posY + 90
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_GERACAO & ") show"
        posY = posY + 70
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.SEQ_OS & ") show"
        posY = posY + 80
'        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.ORIGEM & ") show"
        posY = posY - 45
'        Print #outFN, posX + 300 & " pontos -" & posY & " pontos moveto (" & Reg.EXECUTOR & ") show"
'        Print #outFN, posX + 500 & " pontos -" & posY & " pontos moveto (" & Reg.CAS & ") show"
        posY = posY + 211
        
        'Inscrição, Mátricula
        Print #outFN, posX + 210 & " pontos -" & posY & " pontos moveto (" & Reg.INSCRICAO & ") show"
        Print #outFN, posX + 860 & " pontos -" & posY & " pontos moveto (" & Reg.MATRICULA & ") show"
        Print #outFN, posX + 1330 & " pontos -" & posY & " pontos moveto (" & Reg.PERFIL & ") show"
        posY = posY + 80
        Print #outFN, posX + 290 & " pontos -" & posY & " pontos moveto (" & Reg.ENDERECO & ") show"
        posY = posY + 80
        Print #outFN, posX + 350 & " pontos -" & posY & " pontos moveto (" & Reg.RES & ") show"
        Print #outFN, posX + 510 & " pontos -" & posY & " pontos moveto (" & Reg.COM & ") show"
        Print #outFN, posX + 640 & " pontos -" & posY & " pontos moveto (" & Reg.IND & ") show"
        Print #outFN, posX + 780 & " pontos -" & posY & " pontos moveto (" & Reg.PUB & ") show"
'        Print #outFN, posX + 820 & " pontos -" & posY & " pontos moveto (" & Reg.TOTAL_CATEGORIAS & ") show"
        Print #outFN, posX + 1200 & " pontos -" & posY & " pontos moveto (" & Reg.ULTIMA_ALTERACAO & ") show"
        Print #outFN, posX + 1600 & " pontos -" & posY & " pontos moveto (" & Reg.GRUPO & ") show"
        posY = posY + 80
        Print #outFN, posX + 660 & " pontos -" & posY & " pontos moveto (" & Reg.SITUACAO_AGUA & ") show"
        Print #outFN, posX + 1450 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_MEDIO & ") show"
        posY = posY + 80
        Print #outFN, posX + 660 & " pontos -" & posY & " pontos moveto (" & Reg.SITUACAO_ESGOTO & ") show"
        Print #outFN, posX + 1450 & " pontos -" & posY & " pontos moveto (" & Reg.CONSUMO_FIXO & ") show"
        posY = posY + 45
'        Print #outFN, posX + 320 & " pontos -" & posY & " pontos moveto (" & Reg.OCORRENCIA & ") show"
        posY = posY + 250
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 230 & " pontos -" & posY & " pontos moveto (" & Reg.NOME & ") show"
        posY = posY + 80
        Print #outFN, posX + 280 & " pontos -" & posY & " pontos moveto (" & Reg.CPF_CNPJ & ") show"
        Print #outFN, posX + 1040 & " pontos -" & posY & " pontos moveto (" & Reg.RG & ") show"
        posY = posY + 82
        Print #outFN, posX + 640 & " pontos -" & posY & " pontos moveto (" & Reg.FONE & ") show"
        
        posY = posY + 265
        Print #outFN, posX + 590 & " pontos -" & posY & " pontos moveto (" & Reg.ANORMALIDADE & ") show"
        
        posY = posY + 50
        posY = posY + 840
        'Sequencia e Recuperação
        Print #outFN, "/Helvetica-AP findfont 4.5 scalefont setfont"
        Print #outFN, posX + 130 & " pontos -" & posY & " pontos moveto (Rec: " & Format(wRec, "000000") & "   Seq: " & Reg.SEQ_OS & ") show"
        '**************OUTRA LADO********************
        posY = posY - 2180
        posX = posX + 1600
        Print #outFN, "/Helvetica-Condensed-Bold-AP findfont 8 scalefont setfont"
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.N_OS & ") show"
        posY = posY + 78
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.DATA_GERACAO & ") show"
        posY = posY + 82
        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.SEQ_OS & ") show"
        posY = posY + 90
'        Print #outFN, posX + 1500 & " pontos -" & posY & " pontos moveto (" & Reg.ORIGEM & ") show"
        
        Print #outFN, "showpage"
        Print #outFN, "% Fim-" & Format(CStr(wRec), "00000000000")
        Print #outFN, " "
        
'       If wRec = 1 Then End
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
            
         '----------------------------------------------------------------------------------------------------------------------------------
         'Capturar Registros do arquivo CSV
         '----------------------------------------------------------------------------------------------------------------------------------
                                
           wpv = 0
           For i = 1 To 31
           
               wpv = InStr(1, linha, ";")
               
               If wpv = 0 Then
                  Exit For
               End If
               
               If i = 1 Then RegCSV.C01 = Mid(linha, 1, wpv - 1)
               If i = 2 Then RegCSV.C02 = Mid(linha, 1, wpv - 1)
               If i = 3 Then RegCSV.C03 = Mid(linha, 1, wpv - 1)
               If i = 4 Then RegCSV.C04 = Mid(linha, 1, wpv - 1)
               If i = 5 Then RegCSV.C05 = Mid(linha, 1, wpv - 1)
               If i = 6 Then RegCSV.C06 = Mid(linha, 1, wpv - 1)
               If i = 7 Then RegCSV.C07 = Mid(linha, 1, wpv - 1)
               If i = 8 Then RegCSV.C08 = Mid(linha, 1, wpv - 1)
               If i = 9 Then RegCSV.C09 = Mid(linha, 1, wpv - 1)
               If i = 10 Then RegCSV.C10 = Mid(linha, 1, wpv - 1)
               If i = 11 Then RegCSV.C11 = Mid(linha, 1, wpv - 1)
               If i = 12 Then RegCSV.C12 = Mid(linha, 1, wpv - 1)
               If i = 13 Then RegCSV.C13 = Mid(linha, 1, wpv - 1)
               If i = 14 Then RegCSV.C14 = Mid(linha, 1, wpv - 1)
               If i = 15 Then RegCSV.C15 = Mid(linha, 1, wpv - 1)
               If i = 16 Then RegCSV.C16 = Mid(linha, 1, wpv - 1)
               If i = 17 Then RegCSV.C17 = Mid(linha, 1, wpv - 1)
               If i = 18 Then RegCSV.C18 = Mid(linha, 1, wpv - 1)
               If i = 19 Then RegCSV.C19 = Mid(linha, 1, wpv - 1)
               If i = 20 Then RegCSV.C20 = Mid(linha, 1, wpv - 1)
               If i = 21 Then RegCSV.C21 = Mid(linha, 1, wpv - 1)
               If i = 22 Then RegCSV.C22 = Mid(linha, 1, wpv - 1)
               If i = 23 Then RegCSV.C23 = Mid(linha, 1, wpv - 1)
               If i = 24 Then RegCSV.C24 = Mid(linha, 1, wpv - 1)
               If i = 25 Then RegCSV.C25 = Mid(linha, 1, wpv - 1)
               If i = 26 Then RegCSV.C26 = Mid(linha, 1, wpv - 1)
               If i = 27 Then RegCSV.C27 = Mid(linha, 1, wpv - 1)
               If i = 28 Then RegCSV.C28 = Mid(linha, 1, wpv - 1)
               If i = 29 Then RegCSV.C29 = Mid(linha, 1, wpv - 1)
               If i = 30 Then RegCSV.C30 = Mid(linha, 1, wpv - 1)
               If i = 31 Then RegCSV.C31 = Mid(linha, 1, wpv - 1)
               
               linha = Mid(linha, wpv + 1, 5000)
               
           Next
           
         '----------------------------------------------------------------------------------------------------------------------------------

           With Reg
                .N_OS = RegCSV.C01
                .DATA_GERACAO = RegCSV.C02
                .SEQ_OS = RegCSV.C03
                
                .INSCRICAO = RegCSV.C04
                .MATRICULA = RegCSV.C05
                .PERFIL = RegCSV.C06
                .ENDERECO = RegCSV.C07
                .RES = RegCSV.C08
                .COM = RegCSV.C09
                .IND = RegCSV.C10
                .PUB = RegCSV.C11
                .ULTIMA_ALTERACAO = RegCSV.C12
                .GRUPO = RegCSV.C13
                .SITUACAO_AGUA = RegCSV.C14
                .CONSUMO_MEDIO = RegCSV.C15
                .SITUACAO_ESGOTO = RegCSV.C16
                .CONSUMO_FIXO = RegCSV.C17
                
                .ANORMALIDADE = RegCSV.C26
                
                .NOME = RegCSV.C18
                .CPF_CNPJ = RegCSV.C19
                .RG = RegCSV.C20
                .FONE = RegCSV.C21
           End With
           
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
        Open App.Path & "\InspecaoAnormalidadeInformada.ps" For Input As #200
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
        Print #outFN, "%Aplicação: Fomrmulário de Inspeção de Anormalidade Informada A4"
        Print #outFN,
End Sub
Private Sub LimpaReg()
        With Reg
             .N_OS = ""
             .DATA_GERACAO = ""
             .SEQ_OS = ""
             .INSCRICAO = ""
             .NOME = ""
             .MATRICULA = ""
             .PERFIL = ""
             .ENDERECO = ""
             .RES = ""
             .COM = ""
             .IND = ""
             .PUB = ""
             .ULTIMA_ALTERACAO = ""
             .GRUPO = ""
             .SITUACAO_AGUA = ""
             .CONSUMO_MEDIO = ""
             .SITUACAO_ESGOTO = ""
             .CONSUMO_FIXO = ""
             .NOME = ""
             .CPF_CNPJ = ""
             .RG = ""
             .FONE = ""
        End With
End Sub

