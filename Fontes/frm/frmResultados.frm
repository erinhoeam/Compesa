VERSION 5.00
Begin VB.Form frmResultados 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4995
   Icon            =   "frmResultados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame frameEstatisticas 
      Caption         =   "Verificação"
      Height          =   1725
      Left            =   720
      TabIndex        =   0
      Top             =   1065
      Width           =   3375
      Begin VB.TextBox txArquivo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1230
         Width           =   3015
      End
      Begin VB.TextBox txtTotCli_STL 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   525
         Width           =   1455
      End
      Begin VB.TextBox txtTotReg_STL 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   885
         Width           =   1455
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total de Clientes"
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DATAVOICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   1725
         TabIndex        =   5
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total de Folhas"
         Height          =   330
         Left            =   165
         TabIndex        =   4
         Top             =   885
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   75
      TabIndex        =   8
      Top             =   120
      Width           =   4800
      Begin VB.Frame fraAplicacao 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   4755
         Begin VB.TextBox txtAplicacao 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   135
            TabIndex        =   10
            Top             =   105
            Width           =   4515
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Aplicação"
         Height          =   225
         Left            =   60
         TabIndex        =   11
         Top             =   15
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'Desenvolvido por Aziel R. Pereira Jr
'                 Analista de Sistemas
'Data: Agosto/2002
'Contatos - Email: aziel@hotmail.com
'           Fone: (81) 9979.1972
'Aplicação: Mostrar os Resultados da Geração
'*****************************************************
Private Sub cmdOK_Click()
   Me.Hide
End Sub
Private Sub Form_Load()
        Retangulo Me.hwnd, 30
        Retangulo Frame1.hwnd, 30
        Retangulo fraAplicacao.hwnd, 30
End Sub
