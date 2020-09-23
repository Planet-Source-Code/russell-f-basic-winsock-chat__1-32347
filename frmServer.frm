VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Server"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   1800
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
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' coded by Russell F.

Private Sub Form_Load()
    Winsock1.LocalPort = 400 ' port we listen for connections on
    Winsock1.Listen ' listen...
    txtChat.Text = "Listening..." ' put that in the textbox
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' if enter key is pressed
        Winsock1.SendData "<server> " & txtMsg.Text
        txtChat.Text = txtChat.Text & "<server> " & txtMsg.Text & vbCrLf
        DoEvents
        txtMsg.Text = ""
    End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then Winsock1.Close ' this you should know!
    Winsock1.Accept requestID ' accept the client
    txtChat.Text = txtChat.Text & vbCrLf & "Client connected!" & vbCrLf ' notify the person that a client has connected
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Winsock1.GetData strData
    txtChat.Text = txtChat.Text & strData & vbCrLf
End Sub

