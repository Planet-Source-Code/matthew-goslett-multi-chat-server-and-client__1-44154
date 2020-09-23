VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-Chat Client Example - Matthew Goslett"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckConnection 
      Left            =   7920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6030
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   13203
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab tabbar 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Settings"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNickname"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPort"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblServer"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtNickname"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPort"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtServer"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdConnect"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDisconnect"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Channel #"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstUsers"
      Tab(1).Control(1)=   "txtMessages"
      Tab(1).Control(2)=   "txtCommand"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtCommand 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -74880
         TabIndex        =   2
         Top             =   5040
         Width           =   7935
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "D&isconnect"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         TabIndex        =   10
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   11
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Text            =   "127.0.0.1"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   5
         Text            =   "13234"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtNickname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   4
         Text            =   "Immortality"
         Top             =   2160
         Width           =   2415
      End
      Begin RichTextLib.RichTextBox txtMessages 
         Height          =   4650
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   8202
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0038
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox lstUsers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   -68640
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblServer 
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblPort 
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblNickname 
         Caption         =   "Nickname:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim serverip As String                  'remote server ip
Dim serverport As Integer               'remote server port
Dim nickname As String                  'nickname
Dim data As String                      'received data
Dim splitline() As String               'splits received data by lines
Dim splitdata() As String               'splits received data by spaces
Dim currentchannel As String            'channel you're in

Private Sub SendData(text As String)
If sckConnection.State = 7 Then
    'send data to server
    sckConnection.SendData text & vbCrLf
    'make sure windows follows through with this statement
    DoEvents
End If
End Sub

Private Sub cmdConnect_Click()
If txtServer.text > "" Then
    'store server ip
    serverip = txtServer.text
Else
    'no server ip specified
    StatusBar.Panels(2).text = "You did not enter a server ip"
    Exit Sub
End If
If txtPort.text > "" Then
    'store server port
    serverport = txtPort.text
Else
    'no server port specified
    StatusBar.Panels(2).text = "You did not enter a server port"
    Exit Sub
End If
If txtNickname.text > "" Then
    'store nickname
    nickname = txtNickname.text
Else
    'no nickname specified
    StatusBar.Panels(2).text = "You did not enter a nickname"
    Exit Sub
End If
'disable connect button
cmdConnect.Enabled = False
'enable disconnect button
cmdDisconnect.Enabled = True
'connect to remote server
sckConnection.Connect serverip, serverport
StatusBar.Panels(2).text = "Connecting to " & serverip & ":" & serverport
End Sub

Private Sub cmdDisconnect_Click()
'close connection to server
sckConnection.Close
'hide channel tab
tabbar.TabVisible(1) = False
'disable disconnect button
cmdDisconnect.Enabled = False
'enable connect button
cmdConnect.Enabled = True
StatusBar.Panels(2).text = "Disconnected"
End Sub

Private Sub Form_Load()
StatusBar.Panels(2).text = "Multi-Chat Client"
'hide channel tab
tabbar.TabVisible(1) = False
End Sub

Private Sub sckConnection_Connect()
'show channel tab
tabbar.TabVisible(1) = True
'send nickname to server
SendData "NICKNAME " & nickname
StatusBar.Panels(2).text = "Connected"
End Sub
Private Sub sckConnection_DataArrival(ByVal bytesTotal As Long)
'store received data
sckConnection.GetData data
'split data by lines
splitline = Split(data, vbCrLf, -1, vbTextCompare)
'loop through all lines of data received
For linecount = 0 To UBound(splitline)
    If splitline(linecount) > "" Then
        'split data by spaces
        splitdata = Split(splitline(linecount), " ", -1, vbTextCompare)
        Select Case splitdata(0)
            'JOIN       ( I joined a channel )
            Case "JOIN"
                'set caption of channel tab
                tabbar.TabCaption(1) = "Channel " & splitdata(1)
                'store channel name
                currentchannel = splitdata(1)
                'update messages window
                With txtMessages
                    .SelAlignment = 2
                    .SelColor = vbBlack
                    .SelBold = False
                    .SelFontSize = 13
                    .SelText = "Welcome to" & vbCrLf
                    .SelAlignment = 2
                    .SelBold = True
                    .SelFontSize = 20
                    .SelText = splitdata(1) & vbCrLf & vbCrLf
                    .SelFontSize = 9
                    .SelText = "Topic is: " & Mid(splitline(linecount), InStr(1, splitline(linecount), ":", vbTextCompare) + 1) & vbCrLf & "There are " & splitdata(2) & " users in this channel" & vbCrLf & vbCrLf
                End With
                'add my nickname to user list
                lstUsers.AddItem nickname
            'PART
            Case "PART"
                'set caption of channel tab
                tabbar.TabCaption(1) = "Channel #"
                'unset channel name
                currentchannel = ""
                'clear message window
                txtMessages.text = ""
                'clear command box
                txtCommand.text = ""
                'clear user list
                lstUsers.Clear
            'NICKCHANGE     ( my/a user's nickname has been changed )
            Case "NICKCHANGE"
                If splitdata(1) = nickname Then
                    'my nickname has been changed
                    nickname = splitdata(2)
                    'update messages window
                    With txtMessages
                        .SelAlignment = 0
                        .SelColor = &H8000&
                        .SelText = "* Your nickname is now " & nickname & vbCrLf
                    End With
                Else
                    'a user's nickname has been changed
                End If
            'USERLIST       ( list of nicknames when I join a channel )
            Case "USERLIST"
                'add nickname to user list
                lstUsers.AddItem splitdata(1)
            'USERJOIN       ( a user has joined the channel I'm in )
            Case "USERJOIN"
                'add nickname to user list
                lstUsers.AddItem splitdata(1)
                'update messages window
                With txtMessages
                    .SelAlignment = 0
                    .SelColor = &H8000&
                    .SelText = "* " & splitdata(1) & " has joined " & currentchannel & vbCrLf
                End With
            'USERPART
            Case "USERPART"
                'remove nickname from user list
                For x = 0 To lstUsers.ListCount - 1
                    If lstUsers.List(x) = splitdata(1) Then
                        lstUsers.RemoveItem x
                    End If
                Next x
                'update messages window
                With txtMessages
                    .SelAlignment = 0
                    .SelColor = &H8000&
                    .SelText = "* " & splitdata(1) & " has parted " & currentchannel & vbCrLf
                End With
            'MESSAGE        ( a chat message )
            Case "MESSAGE"
                'update messages window
                With txtMessages
                    .SelAlignment = 0
                    .SelColor = &HC00000
                    .SelText = "<" & splitdata(1) & "> " & Mid(splitline(linecount), InStr(1, splitline(linecount), ":", vbTextCompare) + 1) & vbCrLf
                End With
            'ME             ( an action )
            Case "ME"
                With txtMessages
                    .SelAlignment = 0
                    .SelColor = &H400040
                    .SelText = "* " & splitdata(1) & " " & Mid(splitline(linecount), InStr(1, splitline(linecount), ":", vbTextCompare) + 1) & vbCrLf
                End With
            'QUIT           ( a user has quit the server )
            Case "QUIT"
                'remove nickname from user list
                For x = 0 To lstUsers.ListCount - 1
                    If lstUsers.List(x) = splitdata(1) Then
                        lstUsers.RemoveItem x
                    End If
                Next x
                'update messages window
                With txtMessages
                    .SelAlignment = 0
                    .SelColor = &HC0&
                    .SelText = "* " & splitdata(1) & " quit the server" & vbCrLf
                End With
        End Select
    End If
Next linecount
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'enter key pressed
    If txtCommand.text > "" Then
        If Left(txtCommand.text, 1) = "/" Then
            'command
            splitcmd = Split(txtCommand.text, " ", -1, vbTextCompare)
            Select Case LCase(splitcmd(0))
                '/JOIN
                Case "/join"
                    SendData "JOIN " & nickname & " " & currentchannel & " " & splitcmd(1)
                '/ME
                Case "/me"
                    SendData "ME " & currentchannel & " " & nickname & " :" & Mid(txtCommand.text, 5)
            End Select
        Else
            'chat message
            'send message to server
            SendData "MESSAGE " & currentchannel & " " & nickname & " :" & txtCommand.text
        End If
    End If
    'clear command box
    txtCommand.text = ""
End If
End Sub

Private Sub sckConnection_Close()
'close connection to server
sckConnection.Close
'hide channel tab
tabbar.TabVisible(1) = False
'disable disconnect button
cmdDisconnect.Enabled = False
'enable connect button
cmdConnect.Enabled = True
StatusBar.Panels(2).text = "Disconnected"
End Sub
