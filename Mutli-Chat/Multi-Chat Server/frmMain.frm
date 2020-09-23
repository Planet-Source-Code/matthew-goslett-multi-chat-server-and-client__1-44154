VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-Chat Server Example - Matthew Goslett"
   ClientHeight    =   6030
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
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
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckConnection 
      Index           =   0
      Left            =   120
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
      Top             =   5655
      Width           =   5880
      _ExtentX        =   10372
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
            Object.Width           =   8705
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   8070
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
      TabCaption(0)   =   "Users"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUsers"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstUsers"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Channels"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin MSComctlLib.ListView lstUsers 
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "index"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ip"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "nickname"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblUsers 
         Caption         =   "Users: 0"
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
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim listenport As Integer               'port to listen on
Dim usercount As Integer                'connected user count
Dim intmax As Integer                   'socket count
Dim startdate As String                 'date and time server was started
Dim stopdate As String                  'date and time server was stopped
Dim user(1 To 100) As user_type         'user details
Dim data As String                      'received data
Dim splitline() As String               'splits received data by lines
Dim splitdata() As String               'splits received data by spaces
Dim contentcontrol As Boolean           'activates/deactives swearing blocking
Dim swearword(1 To 100) As String       'swear words to block
Dim channel(1 To 100) As channel_type   'channel details

Private Type user_type
    Index As Integer
    ip As String
    nickname As String
    channel As String
End Type
Private Type channel_type
    name As String
    topic As String
    usercount As Integer
End Type

Private Sub SendData(Index As Integer, text As String)
If sckConnection(Index).State = 7 Then
    'send data to remote user
    sckConnection(Index).SendData text & vbCrLf
    'make sure windows follows through with this statement
    DoEvents
End If
End Sub

Private Sub Form_Load()
'port to listen on
listenport = 13234
'lobby channel name and topic
channel(1).name = "#lobby"
channel(1).topic = "To join a channel, type /join #channelname"
'content control on
contentcontrol = True
'swear words to block
swearword(1) = "fuck"
swearword(2) = "shit"
swearword(3) = "puss"
swearword(4) = "wank"
swearword(5) = "dick"
swearword(6) = "cunt"
StatusBar.Panels(2).text = "Multi-Chat Server"
End Sub

Private Sub cmdStart_click()
'set listen port in winsock control
sckConnection(0).LocalPort = listenport
'listen for connections
sckConnection(0).Listen
'store date and time
startdate = Now
'disable start button
cmdStart.Enabled = False
'enable stop button
cmdStop.Enabled = True
StatusBar.Panels(2).text = "Server started at " & startdate
End Sub

Private Sub cmdStop_click()
'close all user connections
For x = 0 To intmax
    sckConnection(x).Close
Next x
'unload socket controls from memory
For x = 1 To intmax
    Unload sckConnection(x)
Next x
'clear all user details
For x = 1 To 100
    user(x).Index = 0
    user(x).ip = ""
    user(x).nickname = ""
    user(x).channel = ""
Next x
'clear all channel details
For x = 1 To 100
    channel(x).name = ""
    channel(x).topic = ""
    channel(x).usercount = 0
Next x
'clear user list
lstUsers.ListItems.Clear
'clear socket count
intmax = 0
'clear user connected count
usercount = 0
'store date and time
stopdate = Now
'disable stop button
cmdStop.Enabled = False
'enable start button
cmdStart.Enabled = True
'clear user count label
lblUsers.Caption = "Users: 0"
StatusBar.Panels(2).text = "Server stopped at " & stopdate
End Sub

Private Sub sckConnection_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If usercount <= 100 Then
    'increment socket count
    intmax = intmax + 1
    'increment connected user count
    usercount = usercount + 1
    'load new socket for user
    Load sckConnection(intmax)
    'accept user connection
    sckConnection(intmax).Accept requestID
    'store user index and ip
    For x = 1 To 100
        If user(x).Index = 0 Then
            user(x).Index = intmax
            user(x).ip = sckConnection(intmax).RemoteHostIP
            Exit For
        End If
    Next x
    'add user to user list
    With lstUsers
        .ListItems.Add , , intmax
        .ListItems(.ListItems.Count).SubItems(1) = sckConnection(intmax).RemoteHostIP
    End With
    'update user count label
    lblUsers.Caption = "Users: " & usercount
Else
    'server is full
End If
End Sub

Private Sub sckConnection_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'store received data
sckConnection(Index).GetData data
'split data by lines
splitline = Split(data, vbCrLf, -1, vbTextCompare)
'loop through all lines of data received
For linecount = 0 To UBound(splitline)
    If splitline(linecount) > "" Then
        'split data by spaces
        splitdata = Split(splitline(linecount), " ", -1, vbTextCompare)
        Select Case splitdata(0)
            'NICKNAME
            Case "NICKNAME"
                'check if nickname already exists
                For x = 1 To 100
                    If LCase(user(x).nickname) = LCase(splitdata(1)) Then
                        'nickname exists
                        'notify user that his nickname has been changed
                        SendData Index, "NICKCHANGE " & splitdata(1) & " " & "Guest" & intmax
                        'change nickname to guest nickname
                        splitdata(1) = "Guest" & intmax
                    End If
                Next x
                'update user nickname in server user list
                For x = 1 To lstUsers.ListItems.Count
                    If lstUsers.ListItems(x).text = Index Then
                        lstUsers.ListItems(x).SubItems(2) = splitdata(1)
                    End If
                Next x
                'store user's nickname and set his channel to lobby channel
                For x = 1 To 100
                    If user(x).Index = Index Then
                        user(x).nickname = splitdata(1)
                        user(x).channel = channel(1).name
                    End If
                Next x
                'make user join lobby channel and send lobby topic, and usercount
                SendData Index, "JOIN " & channel(1).name & " " & channel(1).usercount & " :" & channel(1).topic
                'send list of user nicknames in lobby channel to new user
                For x = 1 To 100
                    If user(x).nickname > "" Then
                        If user(x).channel = channel(1).name And user(x).Index <> Index Then
                            SendData Index, "USERLIST " & user(x).nickname
                        End If
                    End If
                Next x
                'notify all users in lobby channel that a new user has joined
                For x = 1 To 100
                    If user(x).nickname > "" Then
                        If user(x).channel = channel(1).name And user(x).Index <> Index Then
                            SendData user(x).Index, "USERJOIN " & splitdata(1)
                        End If
                    End If
                Next x
            'MESSAGE
            Case "MESSAGE"
                'store message channel
                message_channel = splitdata(1)
                'store message sender
                message_sender = splitdata(2)
                'store message contents
                message_content = Mid(splitline(linecount), InStr(1, splitline(linecount), ":", vbTextCompare) + 1)
                If contentcontrol = True Then
                    'swear word blocking is on
                    'replace and censor swear words in message content
                    For x = 1 To 100
                        If swearword(x) > "" Then
                            message_content = Replace(message_content, swearword(x), "*****", 1, -1, vbTextCompare)
                        End If
                    Next x
                End If
                'send message to all users in this channel
                For x = 1 To 100
                    If user(x).nickname > "" Then
                        If user(x).channel = message_channel Then
                            SendData user(x).Index, "MESSAGE " & message_sender & " :" & message_content
                        End If
                    End If
                Next x
            'ME
            Case "ME"
                'store message channel
                message_channel = splitdata(1)
                'store message sender
                message_sender = splitdata(2)
                'store message contents
                message_content = Mid(splitline(linecount), InStr(1, splitline(linecount), ":", vbTextCompare) + 1)
                'send action to all users in this channel
                For x = 1 To 100
                    If user(x).nickname > "" Then
                        If user(x).channel = message_channel Then
                            SendData user(x).Index, "ME " & message_sender & " :" & message_content
                        End If
                    End If
                Next x
            'JOIN
            Case "JOIN"
                'make user part previous channel
                SendData Index, "PART"
                'remove channel name from user's details and store new channel name
                For x = 1 To 100
                    If user(x).Index = Index Then
                        user(x).channel = splitdata(3)
                    End If
                Next x
                'notify users in previous channel that this user has parted
                For x = 1 To 100
                    If user(x).nickname > "" Then
                        If user(x).channel = splitdata(2) Then
                            SendData user(x).Index, "USERPART " & splitdata(1)
                        End If
                    End If
                Next x
                'determine if the channel he wants to join already exists
                For x = 1 To 100
                    If channel(x).name = splitdata(3) Then
                        'channel exists so temporarily store channel id
                        channel_exists = x
                        'increment user count for channel
                        channel(x).usercount = channel(x).usercount + 1
                        Exit For
                    End If
                Next x
                If channel_exists > "" Then
                    'channel exists
                    'make user join channel
                    SendData Index, "JOIN " & splitdata(3) & " " & channel(channel_exists).usercount & " :" & channel(channel_exists).topic
                    'send list of user nicknames in the channel he's joining to new user
                    For x = 1 To 100
                        If user(x).nickname > "" Then
                            If user(x).channel = splitdata(3) And user(x).Index <> Index Then
                                SendData Index, "USERLIST " & user(x).nickname
                            End If
                        End If
                    Next x
                    'notify all users in channel that a new user has joined
                    For x = 1 To 100
                        If user(x).nickname > "" Then
                            If user(x).channel = splitdata(3) And user(x).Index <> Index Then
                                SendData user(x).Index, "USERJOIN " & splitdata(1)
                            End If
                        End If
                    Next x
                Else
                    'channel must be created
                    'store channel name and increment channel user count
                    For x = 1 To 100
                        If channel(x).name = "" Then
                            channel(x).name = splitdata(3)
                            channel(x).usercount = 1
                            Exit For
                        End If
                    Next x
                    'make user join channel
                    SendData Index, "JOIN " & splitdata(3) & " 1 :"
                End If
        End Select
    End If
Next linecount
End Sub

Private Sub sckConnection_Close(Index As Integer)
'decrease user connected count
usercount = usercount - 1
'close the connection
sckConnection(Index).Close
'remove user details and temporarily store nickname
For x = 1 To 100
    If user(x).Index = Index Then
        tempnickname = user(x).nickname
        user(x).Index = 0
        user(x).ip = ""
        user(x).nickname = ""
        user(x).channel = ""
    End If
Next x
'remove user from user list
For x = 1 To lstUsers.ListItems.Count
    If lstUsers.ListItems(x).text = Index Then
        lstUsers.ListItems.Remove x
        Exit For
    End If
Next x
'notify all users of user quitting
For x = 1 To 100
    If user(x).Index > 0 Then
        SendData user(x).Index, "QUIT " & tempnickname
    End If
Next x
'update user count label
lblUsers.Caption = "Users: " & usercount
End Sub

