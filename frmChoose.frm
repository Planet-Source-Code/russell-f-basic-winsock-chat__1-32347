VERSION 5.00
Begin VB.Form frmChoose 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server/Client?"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Client"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Server"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' coded by Russell F.

Private Sub Command1_Click()
frmServer.Show
End Sub

Private Sub Command2_Click()
frmClient.Show
End Sub

