VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Connect..."
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   3000
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMsg 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox txtChat 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' coded by Russell F.

Private Sub Command1_Click()
    Winsock1.RemotePort = 400 ' the port that the server is listening on
    Winsock1.Connect txtIP.Text ' lets connect to the IP specified in txtIP
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' If the enter key is pressed
        Winsock1.SendData "<client> " & txtMsg.Text
        txtChat.Text = txtChat.Text & "<client> " & txtMsg.Text & vbCrLf
        DoEvents ' wait until its finished
        txtMsg.Text = "" ' clear the textbox
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String ' strData as string :)
    Winsock1.GetData strData ' lets get the data
    txtChat.Text = txtChat.Text & strData & vbCrLf
End Sub

