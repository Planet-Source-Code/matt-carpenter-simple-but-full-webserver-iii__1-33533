VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Internet Web Server"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Services"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Services"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Full Client Log"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   5655
      Begin RichTextLib.RichTextBox File 
         Height          =   495
         Left            =   3360
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMain.frx":0000
      End
      Begin VB.TextBox txtLog 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Connection Log"
      Height          =   1455
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   2775
      Begin VB.ListBox lstLog 
         Height          =   840
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Hits: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Setup"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
      Begin VB.TextBox txtRootDirectory 
         Height          =   325
         Left            =   1200
         TabIndex        =   7
         Text            =   "C:\"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtPort 
         Height          =   325
         Left            =   600
         TabIndex        =   5
         Text            =   "80"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Root Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   440
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   4080
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox picConnected 
         Height          =   135
         Left            =   600
         ScaleHeight     =   75
         ScaleWidth      =   4395
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
         Begin VB.PictureBox picProgress 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   135
            TabIndex        =   2
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.Image imgClient 
         Height          =   480
         Left            =   5040
         Picture         =   "frmMain.frx":0082
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":04C4
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
cmdStart.Enabled = False
cmdStop.Enabled = True
Frame2.Enabled = False
txtPort.Enabled = False
txtRootDirectory.Enabled = False

'Start the services
Winsock1.LocalPort = txtPort.Text
Winsock1.Listen

End Sub

Private Sub cmdStop_Click()
cmdStart.Enabled = True
cmdStop.Enabled = False
Frame2.Enabled = True
txtPort.Enabled = True
txtRootDirectory.Enabled = True

'Disable Services
Winsock1.Close

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckckosed Then Winsock1.Close
Winsock1.Accept requestID
imgClient.Visible = True
lstLog.AddItem Winsock1.RemoteHostIP & " Connected"
Label3.Caption = "Hits: " & lstLog.ListCount

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Data, vbString, bytesTotal

txtLog.Text = txtLog.Text & Data
aryData = Split(Data, " ", -1, vbBinaryCompare)

If aryData(1) = "/" Then 'User wants homepage
File.LoadFile txtRootDirectory & "index.html"
DoEvents
Winsock1.SendData File.Text
Else
'User is requesting something else
File.LoadFile txtRootDirectory & Right(aryData(1), Len(aryData(1)) - 1)
DoEvents
Winsock1.SendData File.Text

End If

'Enable Progress Counter

imgClient.Visible = True
picConnected.Visible = True
picProgress.Visible = True
picProgress.Width = 1


End Sub

Private Sub Winsock1_SendComplete()
Winsock1.Close
Winsock1.Listen
imgClient.Visible = False
picProgress.Width = 1
picConnected.Visible = False

End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Progress = bytesSent * picConnected.Width / (bytesSent + bytesRemaining)
picProgress.Width = Progress

End Sub
