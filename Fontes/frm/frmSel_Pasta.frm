VERSION 5.00
Begin VB.Form frmSel_Pasta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Selecionar Pasta de Trabalho"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4470
   Icon            =   "frmSel_Pasta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0E0FF&
      Height          =   1065
      Left            =   30
      TabIndex        =   3
      Top             =   3000
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   4455
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "&Selecionar"
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   4155
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0E0FF&
      Height          =   2565
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   4455
   End
End
Attribute VB_Name = "frmSel_Pasta"
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
'Aplicação: Selecionar a pasta de Origem e destino
'*****************************************************
Option Explicit
Private Sub cmdSelecionar_Click()
        Me.Hide
End Sub
Private Sub Drive1_Change()
        Dir1.Path = Drive1.Drive
End Sub
Private Sub Dir1_Click()
        File1.Path = Dir1.Path
End Sub
Private Sub Dir1_Change()
        File1.Path = Dir1.Path
End Sub
