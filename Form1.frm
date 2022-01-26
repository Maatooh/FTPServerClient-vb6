VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FTPClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maatooh FTP Client"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton iHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Connect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock MFTP 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton SFile 
      Caption         =   "Send File"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FTPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FTP(1) As New MaatoohFTP
Public Mode As Boolean 'True Send - False Rec

Private Sub Connect_Click()
With MFTP
.Close
.RemoteHost = "example.com" 'Your Ip.
.RemotePort = 9907
.Connect
End With
Set FTP(0).FTPRec = MFTP
Set FTP(1).FTPSend = MFTP
End Sub

Private Sub iHelp_Click()
FTP(0).Help
End Sub

Private Sub MFTP_DataArrival(ByVal bytesTotal As Long)
Dim xData As String
MFTP.GetData xData
'-----SelectMode
If Not InStr(xData, "Mo1") = 0 Then
Mode = False
xData = Replace(xData, "Mo1", "")
End If
'---Send
If Mode = True Then
FTP(1).PingRec (xData)
Debug.Print FTP(1).TProgressSend
End If
'---Recieved
If Mode = False Then
Call FTP(0).ReceiveFile(xData, App.Path & "\Recibido.wav") 'Any *.*
Debug.Print FTP(0).TProgressRec
End If
End Sub

Private Sub SFile_Click()
If MFTP.State = sckConnected Then
MFTP.SendData "Mo1"
Mode = True
FTP(1).SD = FTP(0).ReadFileMemory(App.Path & "Recibido.wav")  'Any *.*
FTP(1).TransferFile (FTP(1).SD)
End If
End Sub

