VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WAR BOTZ Server"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   Begin VB.ListBox lstID 
      Height          =   3570
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.ListBox lstIP 
      Height          =   3570
      ItemData        =   "frmServer.frx":0000
      Left            =   960
      List            =   "frmServer.frx":0002
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wskConnector 
      Left            =   4680
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.ListBox lstPlayer 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin MSWinsockLib.Winsock wskBot 
      Index           =   0
      Left            =   5160
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Player ID"
      Height          =   195
      Left            =   3360
      TabIndex        =   7
      Top             =   240
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Players IP Adress"
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Socket Connection"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ID_START = 2       'id counter starts at 2

Dim Continue As Boolean          'continue loop checking for a message

Private Sub cmdCopy_Click()      'copy ip to clip board to send for other player to connect to
   Clipboard.Clear
   Clipboard.SetText txtIP.Text
End Sub

Private Sub Form_Load()          'display ip and set ports
   txtIP.Text = wskBot(0).LocalIP
   
   wskConnector.LocalPort = 1
   wskConnector.Listen
   
   wskBot(0).LocalPort = Str(ID_START)
   wskBot(0).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Continue = False
End Sub

Private Sub wskBot_Close(Index As Integer)
   Dim I As Long
   
   '************************* removed in my server
   For I = 0 To lstPlayer.ListCount - 1         'display player info
      If lstPlayer.List(I) = Index Then
         lstPlayer.RemoveItem I
         lstIP.RemoveItem I
         lstID.RemoveItem I

         Exit For
      End If
   Next
   '*************************
   
   If Index = wskBot.ubound - 1 Then      'set socket to listen or unload it from memory
      For I = wskBot.ubound - 1 To wskBot.LBound + 1 Step -1
         If (wskBot(I).State = sckClosing Or wskBot(1).State = sckListening) And wskBot(I - 1).State = sckListening Then
            Unload wskBot(I)
         Else
            Exit For
         End If
      Next
      
      Unload wskBot(wskBot.ubound)
      
      wskBot(wskBot.ubound).Close
      wskBot(wskBot.ubound).Listen
   Else
      wskBot(Index).Close
      wskBot(Index).Listen
   End If
End Sub

Private Sub wskBot_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   'connect and if needed load another socket control
   
   If wskBot(Index).State <> sckClosed Then
      wskBot(Index).Close
      wskBot(Index).Accept requestID
      
      lstPlayer.AddItem Index
      lstIP.AddItem wskBot(Index).RemoteHostIP
      lstID.AddItem wskBot(Index).LocalPort

      If Index = wskBot.ubound Then
         Load wskBot(wskBot.ubound + 1)
         wskBot(wskBot.ubound).LocalPort = Index + ID_START + 1
         wskBot(wskBot.ubound).Listen
      End If
   End If
End Sub

Private Sub wskBot_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   'get message from a player and send to all other players
   Dim message As String   'hold message
   Dim I As Long           'counter
   Dim J As Long
   
   If Continue = False Then      'must have or stack space will run out
      For I = wskBot.LBound To wskBot.ubound - 1
         If wskBot(I).BytesReceived > 0 Then
            wskBot(I).GetData message
            
            For J = wskBot.LBound To wskBot.ubound - 1
               If J <> I And wskBot(J).State = sckConnected Then
                  wskBot(J).SendData message
                                                
                  Continue = True
                  Do While Continue 'wait for socket to finish sending
                     DoEvents
                  Loop
               End If
            Next
         End If
      Next
   End If
End Sub

Private Sub wskBot_SendComplete(Index As Integer)
   Continue = False           'return finished sending
End Sub

Private Sub wskConnector_Close()
   wskConnector.Close         'close socket and listen
   wskConnector.Listen
End Sub

Private Sub wskConnector_ConnectionRequest(ByVal requestID As Long)
   'accept connection and send first available port to connect to
   Dim I As Long
   
   If Continue = False Then
      If wskConnector.State = sckListening Then
         wskConnector.Close
         wskConnector.Accept requestID
         
         For I = wskBot.LBound To wskBot.ubound
            If wskBot(I).State = sckListening Then
               wskConnector.SendData Str(wskBot(I).LocalPort)  'send port player should connect to
               
               Continue = True
               Do While Continue 'wait to finish sending
                  DoEvents
               Loop
               
               Exit For
            End If
         Next
      End If
   End If
End Sub

Private Sub wskConnector_SendComplete()
   Continue = False     'return finished sending
End Sub
