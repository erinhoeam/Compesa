VERSION 5.00
Begin VB.Form frmPrincipal 
   BackColor       =   &H80000013&
   Caption         =   "Gerador de Spool de Impressão COMPESA   "
   ClientHeight    =   7185
   ClientLeft      =   855
   ClientTop       =   975
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmPrincipal.frx":1CCA
   ScaleHeight     =   7185
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Localidade"
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   4230
      Left            =   5400
      TabIndex        =   26
      Top             =   1440
      Width           =   4950
      Begin VB.Frame Frame8 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Width           =   4965
         Begin VB.Frame Frame9 
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            Height          =   1275
            Left            =   0
            TabIndex        =   34
            Top             =   2640
            Width           =   4950
            Begin VB.Frame Frame10 
               BackColor       =   &H80000016&
               BorderStyle     =   0  'None
               Height          =   915
               Left            =   0
               TabIndex        =   35
               Top             =   360
               Width           =   4965
               Begin VB.CheckBox ChkContaD 
                  Caption         =   "Conta D"
                  Height          =   255
                  Left            =   3240
                  TabIndex        =   43
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.CheckBox ChkContaDResp 
                  Caption         =   "Conta D Resp"
                  Height          =   255
                  Left            =   3240
                  TabIndex        =   42
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.CheckBox chkcontaN 
                  Caption         =   "Conta N"
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   41
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.CheckBox chkContanResp 
                  Caption         =   "Conta N Resp"
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   40
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.CheckBox chkBoletoResp 
                  Caption         =   "Boleto Resp"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   39
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.CheckBox chkImprimirVerso 
                  Caption         =   "Imprimir Verso"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   38
                  Top             =   480
                  Width           =   1455
               End
            End
            Begin VB.Label Label7 
               BackColor       =   &H80000003&
               Caption         =   "Configurações"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdPesquisaCArqGerados 
            Height          =   315
            Left            =   4395
            Picture         =   "frmPrincipal.frx":554D
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   315
            Width           =   345
         End
         Begin VB.TextBox txtCaminhoGerado 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   315
            Width           =   4230
         End
         Begin VB.FileListBox LstGerados 
            BackColor       =   &H00FFFFFF&
            Height          =   1560
            Left            =   120
            MultiSelect     =   2  'Extended
            Pattern         =   "*.lst"
            TabIndex        =   31
            Top             =   840
            Width           =   4620
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000016&
            Caption         =   "Diretorio Atual"
            Height          =   255
            Left            =   165
            TabIndex        =   29
            Top             =   75
            Width           =   2415
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000016&
            Caption         =   "Arquivos Gerados (*.ps)"
            Height          =   255
            Left            =   165
            TabIndex        =   28
            Top             =   645
            Width           =   2415
         End
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000003&
         Caption         =   "Arquivos de Gerados (*.ps)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   30
         Top             =   105
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   5400
      TabIndex        =   21
      Top             =   5850
      Width           =   4950
      Begin VB.Frame Frame6 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   4725
         Begin VB.CommandButton cmdProcess 
            Caption         =   "&Dividir Arquivos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1680
            MouseIcon       =   "frmPrincipal.frx":5697
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   165
            Width           =   1485
         End
         Begin VB.CommandButton cmdProcess 
            Caption         =   "De&letar Arquivos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   3255
            MouseIcon       =   "frmPrincipal.frx":57E9
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   165
            Width           =   1380
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir Arquivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   75
            MouseIcon       =   "frmPrincipal.frx":593B
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   165
            Width           =   1530
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   120
      TabIndex        =   17
      Top             =   5850
      Width           =   4950
      Begin VB.Frame Frame2 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   105
         TabIndex        =   18
         Top             =   60
         Width           =   4755
         Begin VB.CommandButton cmdProcesso 
            BackColor       =   &H8000000D&
            Caption         =   "Processar Arquivo"
            Height          =   330
            Left            =   165
            MaskColor       =   &H00C0FFFF&
            TabIndex        =   37
            Top             =   150
            Width           =   1890
         End
         Begin VB.CommandButton cmdProcess 
            Caption         =   "Deleta&r"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   2160
            MouseIcon       =   "frmPrincipal.frx":5A8D
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   150
            Width           =   855
         End
         Begin VB.CommandButton cmdProcess 
            Caption         =   "&Ultimo Resultado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   3120
            MouseIcon       =   "frmPrincipal.frx":5BDF
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   150
            Width           =   1365
         End
      End
   End
   Begin VB.Frame fraImagens 
      Caption         =   "Diretório onde se encontram os arquivos das Imagens TIF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   720
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   9750
      Begin VB.TextBox txtImagens 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   7785
      End
      Begin VB.CommandButton cmdPesquisaImagens 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9240
         TabIndex        =   2
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caminho TIF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.PictureBox StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   6810
      Width           =   10515
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   4230
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4950
      Begin VB.Frame Frame1 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   -15
         TabIndex        =   7
         Top             =   375
         Width           =   4965
         Begin VB.FileListBox LstSpool 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            Left            =   150
            Pattern         =   "*.txt"
            TabIndex        =   12
            Top             =   1485
            Width           =   4620
         End
         Begin VB.ComboBox cmbAplic 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmPrincipal.frx":5D31
            Left            =   150
            List            =   "frmPrincipal.frx":5D33
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Tipo de formatação a ser usada"
            Top             =   285
            Width           =   4635
         End
         Begin VB.TextBox txtCaminhoProcessamento 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   855
            Width           =   4230
         End
         Begin VB.TextBox txQtdReg 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3915
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   3480
            Width           =   855
         End
         Begin VB.CommandButton cmdPesquisaCProcessamento 
            Height          =   315
            Left            =   4425
            Picture         =   "frmPrincipal.frx":5D35
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   855
            Width           =   345
         End
         Begin VB.PictureBox ProgressBarProcess 
            Height          =   360
            Left            =   150
            ScaleHeight     =   300
            ScaleWidth      =   3600
            TabIndex        =   13
            Top             =   3450
            Width           =   3660
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000016&
            Caption         =   "Diretorio Atual"
            Height          =   255
            Left            =   165
            TabIndex        =   16
            Top             =   645
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            Caption         =   "Arquivos de Dados (*.txt)"
            Height          =   300
            Left            =   165
            TabIndex        =   15
            Top             =   1275
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000016&
            Caption         =   "Aplicação"
            Height          =   255
            Left            =   165
            TabIndex        =   14
            Top             =   75
            Width           =   2415
         End
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         Caption         =   "Arquivos de Dados (*.txt)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   6
         Top             =   105
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Sistemap - Consultoria
'Data: Agosto/2008
'Contatos - Email: atendimento@sistemap.com.br
'           Fone: (81) 9979.1972 / (81) 9832.6084
'Aplicação: Gerenciador de Impressão Compesa
'*****************************************************

Option Explicit
'---------------------------- Declaração de Variáveis ---------------------
Private pathProcess As String    'Caminho dos arquivos a serem processados
Private pathGenerated As String  'Caminho dos arquivos gerados

Private Const AppEtiq1      As Integer = 0  'Conta Compesa A4/A5
Private Const AppEtiq2      As Integer = 1  'Conta Compesa Debito Automático A4/A5
Private Const AppEtiq3      As Integer = 2  'Aviso de Corte Compesa Administrativo A4/A5
Private Const AppEtiq4      As Integer = 3  'Aviso de Corte Compesa Boleto A4/A5
Private Const AppEtiq5      As Integer = 4  'Conta Compesa Ficha de Compensação A4
Private Const AppEtiq6      As Integer = 5  'Ordem de Corte Compesa A4/A5
Private Const AppEtiq7      As Integer = 6  'Aviso de Parcelamento em Atraso A4
Private Const AppEtiq8      As Integer = 7  'Carta Instalação Hidrômetro A4/A5
Private Const AppEtiq9      As Integer = 8  'Aviso de Cobrança Compesa A4
Private Const AppEtiq10     As Integer = 9  'Aviso de Cobrança e Perda de Beneficio Compesa A4
Private Const AppEtiq11     As Integer = 10 'Ordem de Substituição Hidrômetro A4/A5
Private Const AppEtiq12     As Integer = 11 'Ordem de Fiscalização Compesa A4
Private Const AppEtiq13     As Integer = 12 'Boletim Cadastral Compesa A4
Private Const AppEtiq14     As Integer = 13 'Ordem de Instalação Hidrômetro Compesa A4/A5
Private Const AppEtiq15     As Integer = 14 'Consumo Inferior a Média Compesa A4/A5
Private Const AppEtiq16     As Integer = 15 'Consumo Superior a Média Compesa A4/A5
Private Const AppEtiq17     As Integer = 16 'Extrato Macromedidor Compesa A4/A5
Private Const AppEtiq18     As Integer = 17 'Extrato Fatura por Responsável A4
Private Const AppEtiq19     As Integer = 18 'Ordem de Fiscalização Inativo A4/A5
Private Const AppEtiq20     As Integer = 19 'Negociação Especial de Débito Em Atraso
Private Const AppEtiq21     As Integer = 20 'Fomulário de Inspeção
Private Const AppEtiq22     As Integer = 21 'Nova Carta de Tarifa Social A4/A5
Private Const AppEtiq23     As Integer = 22 'Ordem de Corte Nova A4
Private Const AppEtiq24     As Integer = 23 'Boletim Cadastral Compesa A4 Novo Com Verso
Private Const AppEtiq25     As Integer = 24 'Aviso de Cobrança Fim do Ano
Private Const AppEtiq26     As Integer = 25 'Declaração Anual de Quitação de Débitos
Private Const AppEtiq27     As Integer = 26 'Ordem de Supressão A4/A5
Private Const AppEtiq28     As Integer = 27 'Carta Urgente
Private Const AppEtiq29     As Integer = 28 'Ordem de Recadastramento Ligação
Private Const AppEtiq30     As Integer = 29 'Inspeção de Esgoto
Private Const AppEtiq31     As Integer = 30 'Inspeção de Anormalidade Informada
Private Const AppEtiq32     As Integer = 31 'Carta Cobrança Desc Encargos
Private Sub Form_Activate()
        LstGerados.Refresh
End Sub
Private Sub LstSpool_Click()
        If Mid(LstSpool.Filename, 1, 7) = "CONTA_N" Or Mid(LstSpool.Filename, 1, 7) = "CONTA_E" Or Mid(LstSpool.Filename, 1, 7) = "CONTA_A" Then
           cmbAplic.ListIndex = (AppEtiq1)
        ElseIf Mid(LstSpool.Filename, 1, 7) = "CONTA_D" Then
           cmbAplic.ListIndex = (AppEtiq2)
        ElseIf Mid(LstSpool.Filename, 1, 29) = "AVISO_DE_CORTE_ADMINISTRATIVO" Or Mid(LstSpool.Filename, 1, 26) = "ORDEM_CORTE_ADMINISTRATIVO" Then
           cmbAplic.ListIndex = (AppEtiq3)
        ElseIf Mid(LstSpool.Filename, 1, 17) = "AVISO_CORTE_GRUPO" Or Mid(LstSpool.Filename, 1, 21) = "AVISO_CORTE_A_REVELIA" Then
           cmbAplic.ListIndex = (AppEtiq4)
        ElseIf Mid(LstSpool.Filename, 1, 8) = "BOLETO_N" Or Mid(LstSpool.Filename, 1, 8) = "BOLETO_E" Or Mid(LstSpool.Filename, 1, 8) = "BOLETO_A" Then
           cmbAplic.ListIndex = (AppEtiq5)
        ElseIf Mid(LstSpool.Filename, 1, 14) = "ORDEM_DE_CORTE" Or Mid(LstSpool.Filename, 1, 18) = "ORDEM_CORTE_FISICO" Then
           cmbAplic.ListIndex = (AppEtiq6)
        ElseIf Mid(LstSpool.Filename, 1, 35) = "EMITIR_CARTA_PARCELAMENTO_EM_ATRASO" Or Mid(LstSpool.Filename, 1, 22) = "CARTAS_DE_PARCELAMENTO" Then
           cmbAplic.ListIndex = (AppEtiq7)
        ElseIf Mid(LstSpool.Filename, 1, 19) = "CARTA_DE_INSTALACAO" Then
           cmbAplic.ListIndex = (AppEtiq8)
        ElseIf Mid(LstSpool.Filename, 1, 20) = "EMITIR_CARTA_CORTADO" Then
           cmbAplic.ListIndex = (AppEtiq9)
        ElseIf Mid(LstSpool.Filename, 1, 26) = "EMITIR_CARTA_TARIFA_SOCIAL" Then
           cmbAplic.ListIndex = (AppEtiq10)
        ElseIf Mid(LstSpool.Filename, 1, 35) = "ORDEM_DE_SUBSTITUICAO_DE_HIDROMETRO" Or Mid(LstSpool.Filename, 1, 15) = "OS_SUBSTITUICAO" Then
           cmbAplic.ListIndex = (AppEtiq11)
        ElseIf Mid(LstSpool.Filename, 1, 21) = "ORDEM_DE_FISCALIZACAO" Then
           cmbAplic.ListIndex = (AppEtiq12)
        ElseIf Mid(LstSpool.Filename, 1, 17) = "BOLETIM_CADASTRAL" Then
           cmbAplic.ListIndex = (AppEtiq13)
        ElseIf Mid(LstSpool.Filename, 1, 33) = "ORDEM_DE_INSTALACAO_DE_HIDROMETRO" Or Mid(LstSpool.Filename, 1, 13) = "OS_INSTALACAO" Then
           cmbAplic.ListIndex = (AppEtiq14)
        ElseIf Mid(LstSpool.Filename, 1, 19) = "CARTA_BAIXO_CONSUMO" Or Mid(LstSpool.Filename, 1, 26) = "CARTA_BOLETO_BAIXO_CONSUMO" Then
           cmbAplic.ListIndex = (AppEtiq15)
        ElseIf Mid(LstSpool.Filename, 1, 19) = "CARTA_CONTA_ESTOURO" Or Mid(LstSpool.Filename, 1, 26) = "CARTA_BOLETO_CONTA_ESTOURO" Then
           cmbAplic.ListIndex = (AppEtiq16)
        ElseIf Mid(LstSpool.Filename, 1, 20) = "EXTRATO_MACROMEDIDOR" Or Mid(LstSpool.Filename, 10, 30) = "extratoConsumoImovelCondominio" Then
           cmbAplic.ListIndex = (AppEtiq17)
        ElseIf Mid(LstSpool.Filename, 1, 27) = "Extrato_Faturas_Responsavel" Then
           cmbAplic.ListIndex = (AppEtiq18)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 29)) = UCase("Ordem_de_FiscalizaCAo_Inativo") Then
           cmbAplic.ListIndex = (AppEtiq19)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 17)) = "CARTA_DE_COBRANCA" Then
           cmbAplic.ListIndex = (AppEtiq20)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 14)) = "ORDEM_INSPECAO" Then
           cmbAplic.ListIndex = (AppEtiq21)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 13)) = "TARIFA SOCIAL" Then
           cmbAplic.ListIndex = (AppEtiq22)
        ElseIf Mid(LstSpool.Filename, 1, 23) = "NOVA_ORDEM_CORTE_FISICO" Then
           cmbAplic.ListIndex = (AppEtiq23)
        ElseIf Mid(LstSpool.Filename, 1, 22) = "NOVO_BOLETIM_CADASTRAL" Then
           cmbAplic.ListIndex = (AppEtiq24)
        ElseIf Mid(LstSpool.Filename, 1, 18) = "CARTA_FINAL_DE_ANO" Then
           cmbAplic.ListIndex = (AppEtiq25)
        ElseIf Mid(LstSpool.Filename, 1, 28) = "DECLARACAO_DE_QUITACAO_ANUAL" Then
           cmbAplic.ListIndex = (AppEtiq26)
        ElseIf Mid(LstSpool.Filename, 1, 18) = "ORDEM_DE_SUPRESSAO" Then
           cmbAplic.ListIndex = (AppEtiq27)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 13)) = "CARTA_URGENTE" Then
           cmbAplic.ListIndex = (AppEtiq28)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 21)) = "ORDEM_RECADASTRAMENTO" Then
           cmbAplic.ListIndex = (AppEtiq29)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 6)) = "ESGOTO" Then
           cmbAplic.ListIndex = (AppEtiq30)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 24)) = "OS_INSPECAO_ANORMALIDADE" Then
           cmbAplic.ListIndex = (AppEtiq31)
        ElseIf UCase(Mid(LstSpool.Filename, 1, 5)) = "CT_CB" Then
           cmbAplic.ListIndex = (AppEtiq32)
        Else
           MsgBox "Aplicação não Localizada!", vbInformation + vbOKOnly, "AP SERVIÇOS DE INFORMÁTICA"
        End If
End Sub
'------------------------------------- Código -------------------------------------
Private Sub mnAtualizar_Click()
        LstSpool.Path = pathProcess
        LstSpool.Pattern = "*.txt"
        LstSpool.Refresh
        LstGerados.Path = pathGenerated
        LstGerados.Pattern = "*.ps"
        LstGerados.Refresh
End Sub
Private Sub mnSair_Click()
   Unload Me
End Sub
Private Sub cmdProcesso_Click()
        Dim wcont_for As Integer
        Dim campanha As String
        Dim resposta
        cmdProcesso.Enabled = False
        MakeTransparent Me.hwnd, 150
   
        'Processa o arquivo selecionado da lista dos .txt
        For wcont_for = 0 To LstSpool.ListCount - 1
            If LstSpool.Selected(wcont_for) = True Then
               Processar pathProcess & LstSpool.List(wcont_for)
               Exit For
               
            End If
        Next
        MakeOpaque Me.hwnd
        cmdProcesso.Enabled = True
End Sub
Private Sub cmdPesquisaCProcessamento_Click()
        Dim fName As String
        fName = App.Path & "\ConfigSec.ini"
        
        frmSel_Pasta.Dir1 = txtCaminhoProcessamento.Text
        frmSel_Pasta.Show vbModal
        If Len(frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex)) > 3 Then
           txtCaminhoProcessamento.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex) & "\"
           txtCaminhoGerado.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex) & "\"
        Else
           txtCaminhoProcessamento.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex)
           txtCaminhoGerado.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex)
        End If
        
        writeINI "CONFIG", "Processamento", txtCaminhoProcessamento.Text, fName
        writeINI "CONFIG", "Gerados", txtCaminhoGerado.Text, fName
        
        pathProcess = txtCaminhoProcessamento.Text
        pathGenerated = txtCaminhoGerado.Text
        
        LstSpool.Path = pathProcess
        LstSpool.Pattern = "*.txt"
        LstSpool.Refresh
        
        LstGerados.Path = pathGenerated
        LstGerados.Pattern = "*.ps"
        LstGerados.Refresh
 
End Sub

Private Sub cmdPesquisaCArqGerados_Click()
   Dim fName As String
   fName = App.Path & "\ConfigSec.ini"
   
   frmSel_Pasta.Dir1 = txtCaminhoGerado.Text
   frmSel_Pasta.Show vbModal
   If Len(frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex)) > 3 Then
      txtCaminhoGerado.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex) & "\"
   Else
      txtCaminhoGerado.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex)
   End If
   writeINI "CONFIG", "Gerados", txtCaminhoGerado.Text, fName
   pathGenerated = txtCaminhoGerado.Text
   LstGerados.Path = pathGenerated
   LstGerados.Pattern = "*.lst;*.ps"
   LstGerados.Refresh
End Sub

Private Sub cmdPesquisaImagens_Click()
   Dim fName As String
   fName = App.Path & "\ConfigSec.ini"

   frmSel_Pasta.Dir1 = txtImagens.Text
   frmSel_Pasta.Show vbModal
   If Len(frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex)) > 3 Then
      txtImagens.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex) & "\"
   Else
      txtImagens.Text = frmSel_Pasta.Dir1.List(frmSel_Pasta.Dir1.ListIndex)
   End If
   
   
   fraImagens.Visible = False
   Me.Height = 7685
End Sub

Private Sub cmdImprimir_Click()
Dim wcont_for As Integer

   For wcont_for = 0 To frmPrincipal.LstGerados.ListCount - 1
      If frmPrincipal.LstGerados.Selected(wcont_for) = True Then
         'frmImprimir.printFile pathGenerated & LstGerados.List(wcont_for)
         'frmImprimir.Show
      End If
   Next
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ShiftKey
   ShiftKey = Shift And 7
   
   If KeyCode = vbKeyI And ShiftKey = vbShiftMask Then
      fraImagens.Top = 1680
      fraImagens.Visible = True
   End If
'*******************************************************
   If KeyCode = vbKeyEscape Then
      Unload Me
   End If
'*******************************************************
End Sub
Private Sub Processar(iFile As String)
       Dim inLine As String    'Linha lida
       Dim fNIn As Integer     'Identificadores do arquivo de entrada
       Dim aplic As String
    '-----------------------------------------------------------------------------------------
       frmPrincipal.MousePointer = vbHourglass
       'StatusBar.Panels(1) = "Processando arquivo " & iFile & "."
       
       fNIn = FreeFile
       Open iFile For Random As #fNIn Len = 1
       
       'Ler primeira linha do arquivo de entrada
       ler fNIn, inLine
    
    '*****************************************************************
    '*********** Processamento Extratos Compesa *********************
    '*****************************************************************
       If cmbAplic.ListIndex = AppEtiq1 Then
    
          'StatusBar.Panels(2) = "Conta Compesa A4/A5"
          MousePointer = vbHourglass
    
          Close fNIn
    
          Dim objContaCOMPESA As clsConta_Compesa_A5
          Set objContaCOMPESA = New clsConta_Compesa_A5
          objContaCOMPESA.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq2 Then

          'StatusBar.Panels(2) = "Conta Compesa Debito Automático A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objContaCOMPESA_Debito_Automatico As clsConta_Deb_Aut_A5
          Set objContaCOMPESA_Debito_Automatico = New clsConta_Deb_Aut_A5
          objContaCOMPESA_Debito_Automatico.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq3 Then

          'StatusBar.Panels(2) = "Aviso de Corte Compesa Administrativo A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objAvisoDeCorteAdministrativoCompesa As clsAvisodeCorte_Admin_A5
          Set objAvisoDeCorteAdministrativoCompesa = New clsAvisodeCorte_Admin_A5
          objAvisoDeCorteAdministrativoCompesa.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq4 Then

          'StatusBar.Panels(2) = "Aviso de Corte Compesa Boleto A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objAvisoDeCorteBoletoCompesa As clsAvisodeCorte_Boleto_A5
          Set objAvisoDeCorteBoletoCompesa = New clsAvisodeCorte_Boleto_A5
          objAvisoDeCorteBoletoCompesa.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
       ElseIf cmbAplic.ListIndex = AppEtiq5 Then

          'StatusBar.Panels(2) = "Conta Compesa Ficha de Compensação A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objContaCompesaFichaDeCompensacao As clsConta_Boleto_Compesa_A4
          Set objContaCompesaFichaDeCompensacao = New clsConta_Boleto_Compesa_A4
          objContaCompesaFichaDeCompensacao.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq6 Then

          'StatusBar.Panels(2) = "Ordem de Corte Compesa A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objOrdemDeCortecompesa_A5 As clsOrdem_de_Corte_Compesa_A5
          Set objOrdemDeCortecompesa_A5 = New clsOrdem_de_Corte_Compesa_A5
          objOrdemDeCortecompesa_A5.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq7 Then

          'StatusBar.Panels(2) = "Aviso de Parcelamento em Atraso A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objAvisoParcelamentoEmAtrasoA4 As clsAviso_de_Parc_em_Atraso_A4
          Set objAvisoParcelamentoEmAtrasoA4 = New clsAviso_de_Parc_em_Atraso_A4
          objAvisoParcelamentoEmAtrasoA4.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq8 Then

          'StatusBar.Panels(2) = "Carta de Instalação de Hidrômetro Compesa A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objOrdemDeInstalacaoHidrometro As clsCarta_Ins_Hid_A5
          Set objOrdemDeInstalacaoHidrometro = New clsCarta_Ins_Hid_A5
          objOrdemDeInstalacaoHidrometro.DoProcess iFile, pathProcess, pathGenerated


       ElseIf cmbAplic.ListIndex = AppEtiq9 Then

          'StatusBar.Panels(2) = "Aviso de Cobrança Compesa A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objAvisoDeCobrancaCompesaA4 As clsAviso_de_Cobranca_Compesa_A4
          Set objAvisoDeCobrancaCompesaA4 = New clsAviso_de_Cobranca_Compesa_A4
          objAvisoDeCobrancaCompesaA4.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq10 Then

          'StatusBar.Panels(2) = "Aviso de Cobrança e Perda de Beneficio Compesa A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objAvisoDeCobrancaPerdaDeBeneficioCompesaA4 As clsAvisodeCob_e_Perda_Benef_A4
          Set objAvisoDeCobrancaPerdaDeBeneficioCompesaA4 = New clsAvisodeCob_e_Perda_Benef_A4
          objAvisoDeCobrancaPerdaDeBeneficioCompesaA4.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq11 Then

          'StatusBar.Panels(2) = "Ordem de Substituição de Hidrômetro Compesa A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objOrdemDeSubstituicaoHidrometro As clsOrdem_Substituicao_Hid_A5
          Set objOrdemDeSubstituicaoHidrometro = New clsOrdem_Substituicao_Hid_A5
          objOrdemDeSubstituicaoHidrometro.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq12 Then

          'StatusBar.Panels(2) = "Ordem de Fiscalização Compesa A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objOrdemDeFiscalizacao As clsOrdem_de_Fisc_A4
          Set objOrdemDeFiscalizacao = New clsOrdem_de_Fisc_A4
          objOrdemDeFiscalizacao.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq13 Then

          'StatusBar.Panels(2) = "Boletim Cadastral Compesa A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objBoletimCadastral_Novo As clsBoletim_Cad_Compesa_A4_Novo
          Set objBoletimCadastral_Novo = New clsBoletim_Cad_Compesa_A4_Novo
          objBoletimCadastral_Novo.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq14 Then

          'StatusBar.Panels(2) = "Ordem de Instalação Hidrômetro Compesa A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objordemInstalacaoHidrometro As clsOrdem_Inst_Hidrometro_A5
          Set objordemInstalacaoHidrometro = New clsOrdem_Inst_Hidrometro_A5
          objordemInstalacaoHidrometro.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq15 Then

          'StatusBar.Panels(2) = "Consumo Inferior a Média"
          MousePointer = vbHourglass

          Close fNIn

          Dim objConsumoInferiorMedia As clsConsumoInferior_Media_A5
          Set objConsumoInferiorMedia = New clsConsumoInferior_Media_A5
          objConsumoInferiorMedia.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq16 Then

          'StatusBar.Panels(2) = "Consumo Superior a Média"
          MousePointer = vbHourglass

          Close fNIn

          Dim objConsumoSuperiorMedia As clsConsumoSuperior_Media_A5
          Set objConsumoSuperiorMedia = New clsConsumoSuperior_Media_A5
          objConsumoSuperiorMedia.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq17 Then

          'StatusBar.Panels(2) = "Extrato Macromedidor"
          MousePointer = vbHourglass

          Close fNIn

          Dim objExtratoMacromedidor As clsExtrato_Macromedidor_A5
          Set objExtratoMacromedidor = New clsExtrato_Macromedidor_A5
          objExtratoMacromedidor.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq18 Then

          'StatusBar.Panels(2) = "Extrato Fatura por Responsável Compesa A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objExtratoFaturaResponsavel As clsExt_Fat_Resp
          Set objExtratoFaturaResponsavel = New clsExt_Fat_Resp
          objExtratoFaturaResponsavel.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq19 Then

          'StatusBar.Panels(2) = "Ordem de Fiscalizacao Inativos A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objOrdeFiscaInat As clsOrdem_Fisc_Inat_Compesa_A5
          Set objOrdeFiscaInat = New clsOrdem_Fisc_Inat_Compesa_A5
          objOrdeFiscaInat.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq20 Then

          'StatusBar.Panels(2) = "Negociação Especial de Débito Em Atraso"
          MousePointer = vbHourglass

          Close fNIn

          Dim objNegociacaoEspecialDebito As clsNeg_Especial_Debito_Avista
          Set objNegociacaoEspecialDebito = New clsNeg_Especial_Debito_Avista
          objNegociacaoEspecialDebito.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq21 Then

          'StatusBar.Panels(2) = "Formulário de Inspeção Compesa A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objFomrularioInpecao As clsForm_de_Inspec_A4
          Set objFomrularioInpecao = New clsForm_de_Inspec_A4
          objFomrularioInpecao.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq22 Then

          'StatusBar.Panels(2) = "Nova Carta de Tarifa Social A4/A5"
          MousePointer = vbHourglass

          Close fNIn

          Dim objTarifaSocial As clsTarifa_Social_Compesa_A5
          Set objTarifaSocial = New clsTarifa_Social_Compesa_A5
          objTarifaSocial.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq23 Then

          'StatusBar.Panels(2) = "Ordem de Corte Compesa A4 Nova"
          MousePointer = vbHourglass

          Close fNIn

          Dim objOrdemDeCortecompesa_A4_New As clsOrdem_Corte_New_Compesa_A4
          Set objOrdemDeCortecompesa_A4_New = New clsOrdem_Corte_New_Compesa_A4
          objOrdemDeCortecompesa_A4_New.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq24 Then

          'StatusBar.Panels(2) = "Boletim Cadastral Compesa A4 Novo Verso"
          MousePointer = vbHourglass

          Close fNIn

          Dim objBoletimCadastral_Novo_Verso As clsBolet_Cad_Compesa_A4_NovoVer
          Set objBoletimCadastral_Novo_Verso = New clsBolet_Cad_Compesa_A4_NovoVer
          objBoletimCadastral_Novo_Verso.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq25 Then

          'StatusBar.Panels(2) = "Boletim Cadastral Compesa A4 Novo Verso"
          MousePointer = vbHourglass

          Close fNIn

          Dim objAviso_Cob_FimAno_Compesa_A4 As clsAviso_Cob_FimAno_Compesa_A4
          Set objAviso_Cob_FimAno_Compesa_A4 = New clsAviso_Cob_FimAno_Compesa_A4
          objAviso_Cob_FimAno_Compesa_A4.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq26 Then

          'StatusBar.Panels(2) = "Declaração de Quitação Anula de Débito"
          MousePointer = vbHourglass

          Close fNIn

          Dim objDeclaracaoQuitAnual As clsDeclaracaoAnualDebito
          Set objDeclaracaoQuitAnual = New clsDeclaracaoAnualDebito
          objDeclaracaoQuitAnual.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq27 Then

          'StatusBar.Panels(2) = "Ordem de Supressão"
          MousePointer = vbHourglass

          Close fNIn

          Dim objOrdemSupressao As clsOrdemDeSupressao
          Set objOrdemSupressao = New clsOrdemDeSupressao
          objOrdemSupressao.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq28 Then

          'StatusBar.Panels(2) = "Carta Urgente A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objCartaUrgenteA4 As clsCartaUrgente_A4
          Set objCartaUrgenteA4 = New clsCartaUrgente_A4
          objCartaUrgenteA4.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault
          
       ElseIf cmbAplic.ListIndex = AppEtiq29 Then

          'StatusBar.Panels(2) = "Ordem de Racadastramento de Ligação A4"
          MousePointer = vbHourglass

          Close fNIn

          Dim objclsRecadastramentoLigacaoA4 As clsRecadastramentoLigacaoA4
          Set objclsRecadastramentoLigacaoA4 = New clsRecadastramentoLigacaoA4
          objclsRecadastramentoLigacaoA4.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq30 Then

          'StatusBar.Panels(2) = "Inspeção de Esgoto"
          MousePointer = vbHourglass

          Close fNIn

          Dim objclsInspecaoDeEsgoto As clsInspecaoDeEsgoto
          Set objclsInspecaoDeEsgoto = New clsInspecaoDeEsgoto
          objclsInspecaoDeEsgoto.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq31 Then

          'StatusBar.Panels(2) = "Inspeção de Anormalidade Informada"
          MousePointer = vbHourglass

          Close fNIn

          Dim objclsInspecaoAnormalidadeInformada As clsInsp_Anor_Inf_A4
          Set objclsInspecaoAnormalidadeInformada = New clsInsp_Anor_Inf_A4
          objclsInspecaoAnormalidadeInformada.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       ElseIf cmbAplic.ListIndex = AppEtiq32 Then

          'StatusBar.Panels(2) = "Carta Cobranca Desc. Encargos"
          MousePointer = vbHourglass

          Close fNIn

          Dim objCarta_cobranca_desc_enc_A4 As clsCarta_cobranca_desc_enc_A4
          Set objCarta_cobranca_desc_enc_A4 = New clsCarta_cobranca_desc_enc_A4
          objCarta_cobranca_desc_enc_A4.DoProcess iFile, pathProcess, pathGenerated

          frmResultados.Show
          LstGerados.Refresh
          MousePointer = vbDefault

       Else
          If cmbAplic.ListIndex = -1 Then
             MsgBox "Selecione qual o tipo de aplicação a ser processada!", vbInformation
          Else
             MsgBox "Formato de arquivo inválido, ou o tipo da aplicação selecionada está errado." & vbCrLf & "Verifique e tente novamente.", vbCritical, "ERRO"
          End If
          Close fNIn
          frmPrincipal.MousePointer = vbNormal
       End If
End Sub
Private Sub Form_Load()
        
        Dim wdrRecordset As New ADODB.Recordset
        wStringConexao = "DRIVER={MySQL ODBC 3.51 Driver};server=201.73.10.41;database=sistema;uid=sistemap;"

'        wdrRecordset.CursorLocation = adUseClient
'        wdrRecordset.Open "select * from bloq;", wStringConexao, adOpenForwardOnly, adLockOptimistic
'        If wdrRecordset("status") = "S" Then
'           MsgBox "Critical Error!", vbOKOnly + vbCritical, "RUN TIME ERROR"
'           End
'        End If
        
        Dim fName As String
        Dim i As Integer, x As Integer
        Dim prn As Printer
        
        Me.Height = 7695
        Me.Width = 10635
        
        'Dim lR As Long
        'lR = SetTopMostWindow(Me.hwnd, True)

       'Carregar
       fName = App.Path & "\ConfigSec.ini"
       If Manip_Arq(EXISTEARQ, fName) Then
          pathProcess = readINI("CONFIG", "Processamento", fName)
          txtCaminhoProcessamento.Text = pathProcess

          pathGenerated = readINI("CONFIG", "Gerados", fName)
          txtCaminhoGerado.Text = pathGenerated
      
          If Not Manip_Arq(EXISTEPASTA, pathProcess) Then
             pathProcess = App.Path & "\"
             writeINI "CONFIG", "Processamento", pathProcess, App.Path & "\ConfigSec.ini"
          End If

          If Not Manip_Arq(EXISTEPASTA, pathGenerated) Then
             pathGenerated = App.Path & "\"
             writeINI "CONFIG", "Gerados", pathGenerated, App.Path & "\ConfigSec.ini"
          End If

      Else
          Dim fso As Object
          Dim op As Object
          Set fso = CreateObject("Scripting.FileSystemObject")
          Set op = fso.CreateTextFile(fName, True)
          op.WriteLine ("[CONFIG]")
          op.WriteLine ("Processamento=" & App.Path & "\")
          op.WriteLine ("Gerados=" & App.Path & "\")
          op.Close

          pathProcess = readINI("CONFIG", "Processamento", fName)
          txtCaminhoProcessamento.Text = pathProcess

          pathGenerated = readINI("CONFIG", "Gerados", fName)
          txtCaminhoGerado.Text = pathGenerated
     End If

     cmbAplic.List(AppEtiq1) = "Conta Compesa A4/A5"
     cmbAplic.List(AppEtiq2) = "Conta Compesa Débito Automático A4/A5"
     cmbAplic.List(AppEtiq3) = "Aviso de Corte Administrativo Compesa A4/A5"
     cmbAplic.List(AppEtiq4) = "Aviso de Corte Boleto Compesa A4/A5"
     cmbAplic.List(AppEtiq5) = "Conta Compesa Ficha de Compensação A4"
     cmbAplic.List(AppEtiq6) = "Ordem de Corte Compesa A4/A5"
     cmbAplic.List(AppEtiq7) = "Aviso de Parcelamento em Atraso A4"
     cmbAplic.List(AppEtiq8) = "Carta de Instalação de Hidrômetro A4/A5"
     cmbAplic.List(AppEtiq9) = "Aviso de Cobrança Compesa A4"
     cmbAplic.List(AppEtiq10) = "Aviso de Cobrança  e Perda de Beneficio Compesa A4"
     cmbAplic.List(AppEtiq11) = "Ordem de Substituição de Hidrômetro A4/A5"
     cmbAplic.List(AppEtiq12) = "Ordem de Fiscalização Compesa A4"
     cmbAplic.List(AppEtiq13) = "Boletim Cadastral Compesa A4"
     cmbAplic.List(AppEtiq14) = "Ordem de Instalação de Hidrômetro A4/A5"
     cmbAplic.List(AppEtiq15) = "Consumo Inferior a Média A4/A5"
     cmbAplic.List(AppEtiq16) = "Consumo Superior a Média A4/A5"
     cmbAplic.List(AppEtiq17) = "Extrato Macromedidor A4/A5"
     cmbAplic.List(AppEtiq18) = "Extrato Fatura por Responsável Compesa A4"
     cmbAplic.List(AppEtiq19) = "Ordem de Fiscalização Inativos A4/A5"
     cmbAplic.List(AppEtiq20) = "Negociação Especial de Débito Em Atraso"
     cmbAplic.List(AppEtiq21) = "Formulário de Inspeção Compesa A4"
     cmbAplic.List(AppEtiq22) = "Nova Carta de Tarifa Social A4/A5"
     cmbAplic.List(AppEtiq23) = "Ordem de Corte Nova A4"
     cmbAplic.List(AppEtiq24) = "Boletim Cadastral Compesa A4 Novo Com Verso"
     cmbAplic.List(AppEtiq25) = "Aviso de Cobrança Final do Ano"
     cmbAplic.List(AppEtiq26) = "Declaração de Quitação Anual de Débito"
     cmbAplic.List(AppEtiq27) = "Ordem de Supressão A4/A5"
     cmbAplic.List(AppEtiq28) = "Carta Urgente"
     cmbAplic.List(AppEtiq29) = "Recadastramento Ligação A4"
     cmbAplic.List(AppEtiq30) = "Inspeção de Esgoto"
     cmbAplic.List(AppEtiq31) = "Inspeção de Anormalidade Informada A4"
     cmbAplic.List(AppEtiq32) = "Carta Cobrança Desc. Encargos"
     cmbAplic.ListIndex = 0
     
     Retangulo Me.hwnd, 30
     
     Retangulo Frame1.hwnd, 30
     Retangulo Frame2.hwnd, 30
     Retangulo Frame3.hwnd, 30
     Retangulo Frame4.hwnd, 30
     Retangulo Frame5.hwnd, 30
     Retangulo Frame6.hwnd, 30
     Retangulo Frame7.hwnd, 30
     Retangulo Frame8.hwnd, 30
     Retangulo Frame9.hwnd, 30
     Retangulo Frame10.hwnd, 30
     
     Retangulo cmdProcesso.hwnd, 30

     mnAtualizar_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
        End
End Sub
Private Sub LstGerados_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim wcont_for As Integer
        If KeyCode = vbKeyDelete Then
           For wcont_for = 0 To LstGerados.ListCount - 1
               If LstGerados.Selected(wcont_for) = True Then
                  If Not Manip_Arq(DELETAR, pathGenerated & LstGerados.List(wcont_for)) Then
                     MsgBox "Erro ao tentar deletar o arquivo " & LstGerados.List(wcont_for) & "."
                  End If
                  Exit For
               End If
           Next
           LstGerados.Refresh
        End If
End Sub
Private Sub LstSpool_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim wcont_for As Integer
        If KeyCode = vbKeyDelete Then
           For wcont_for = 0 To LstSpool.ListCount - 1
               If LstSpool.Selected(wcont_for) = True Then
                  If Not Manip_Arq(DELETAR, pathGenerated & LstSpool.List(wcont_for)) Then
                     MsgBox "Erro ao tentar deletar o arquivo " & LstSpool.List(wcont_for) & "."
                  End If
                  Exit For
               End If
           Next
           LstSpool.Refresh
        End If
End Sub

