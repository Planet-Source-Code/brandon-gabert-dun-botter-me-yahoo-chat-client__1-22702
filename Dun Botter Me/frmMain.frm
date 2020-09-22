VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dun Botter Me"
   ClientHeight    =   6870
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   5520
      Width           =   7095
   End
   Begin VB.ListBox lstCustomRooms 
      Height          =   1230
      Left            =   0
      TabIndex        =   17
      Top             =   3960
      Width           =   3375
   End
   Begin VB.ListBox lstSubRooms 
      Height          =   2205
      Left            =   4920
      TabIndex        =   16
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ListBox lstMainRooms 
      Height          =   2205
      Left            =   2400
      TabIndex        =   15
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ListBox lstMainCats 
      Height          =   2205
      ItemData        =   "frmMain.frx":0000
      Left            =   0
      List            =   "frmMain.frx":0040
      TabIndex        =   14
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdSay 
      Caption         =   "Say"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtSpeak 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Top             =   6600
      Width           =   5415
   End
   Begin VB.CommandButton cmdJoinRoom 
      Caption         =   "Join Room"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdAuth 
      Caption         =   "Authorize"
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wskChat 
      Left            =   6600
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wskAuth 
      Left            =   6600
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskRooms 
      Left            =   6600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label11 
      Caption         =   $"frmMain.frx":0184
      Height          =   1215
      Left            =   3480
      TabIndex        =   24
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Chat Here"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   5280
      Width           =   7095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "SubRooms"
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Main Rooms"
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Main Categories"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Custom Rooms"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Speak"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label lblRoom 
      Caption         =   "none"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Current Room"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Connection Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Caption         =   "Not Connected"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server Options"
      Begin VB.Menu mnuAuthServer 
         Caption         =   "&Auth Server"
      End
      Begin VB.Menu mnuChatServer 
         Caption         =   "&Chat Server"
      End
      Begin VB.Menu mnuPort 
         Caption         =   "Chat &Port"
      End
   End
   Begin VB.Menu mnuHelpM 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private thePacket As String
Private lastCommand As String
Private currentRoom As String
Private areWeConnected As String
Private Room As String
Private subRoom As String
Private roomsVariable() As String
Private subRoomsVariable() As String
Private customRoomsVariable() As String
Private roomsMode As Boolean
Private AuthServer As String
Private ChatServer As String
Private ChatPort As Integer
Private UserName As String
Private PassWord As String
Private Cookie As String

'Yahoo Version Number
Private cYahooVer As String
'Packet Types
Private cLogin As String
Private cLogout As String
Private cRoomEnter As String
Private cRoomLeave  As String
Private cInvitations  As String
Private cAwayBack  As String
Private cSpeak As String
Private cThink As String
Private cEmote As String
Private cAdvertisements As String
Private cPrivMesg As String
Private cBuddyListMesg As String
Private cYahooMail As String
Private cYahooMesg As String
Private cGrafiti As String

Private sLogin As String
Private sLogout As String
Private sRoomEnter As String
Private sRoomLeave  As String
Private sInvitations  As String
Private sAwayBack  As String
Private sSpeak As String
Private sThink As String
Private sEmote As String
Private sAdvertisements As String
Private sPrivMesg As String
Private sBuddyListMesg As String
Private sYahooMail As String
Private sYahooMesg As String
Private SGrafiti As String

'declare sleep function
Private Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)

Private Sub cmdAuth_Click()
    If UserName = vbNullString Or _
       PassWord = vbNullString Then
        MsgBox "Please supply a username and password"
        Exit Sub
    End If
    getCookie
End Sub

Private Sub cmdConnect_Click()
    If wskChat.State <> sckClosed Then wskChat.Close
    wskChat.RemoteHost = ChatServer
    wskChat.RemotePort = ChatPort
    wskChat.Connect
    Dim fatalErrorAK47 As Integer
    While wskChat.State <> sckConnected
        fatalErrorAK47 = fatalErrorAK47 + 1
        DoEvents
        If fatalErrorAK47 = 30000 Then
            MsgBox "Could not connect to " & ChatServer
        End If
    Wend
    
    'create the login packet
    Dim loginPacket As String
    Dim loginPacketInfo As String
    loginPacketInfo = UserName & Chr$(1) & Cookie
    loginPacket = createPacket(cLogin, loginPacketInfo)
    wskChat.SendData loginPacket
End Sub

Private Sub cmdJoinRoom_Click()
    Dim theroom As String
    theroom = InputBox("Enter the room you wish to join", "Join Room", "Movies:1")
    joinRoom (theroom)
End Sub

Private Sub cmdSay_Click()
    If txtSpeak.Text = vbNullString Then
        MsgBox "You need something to say"
        Exit Sub
    End If
    Speak (txtSpeak.Text)
End Sub

Private Sub Form_Load()
    'Yahoo Version Number
    cYahooVer = Chr$(0) & Chr$(0) & Chr$(1) & Chr$(107)
    'Packet Types
    cLogin = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(1)
    cLogout = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(2)
    cRoomEnter = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(113)
    cRoomLeave = Chr$(0) & Chr$(0) & Chr$(1) & Chr$(2)
    cInvitations = Chr$(0) & Chr$(0) & Chr$(1) & Chr$(7)
    cAwayBack = Chr$(0) & Chr$(0) & Chr$(2) & Chr$(1)
    cSpeak = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(65)
    cThink = Chr$(0) & Chr$(0) & Chr$(4) & Chr$(2)
    cEmote = Chr$(0) & Chr$(0) & Chr$(4) & Chr$(3)
    cAdvertisements = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(68)
    cPrivMesg = Chr$(0) & Chr$(0) & Chr$(4) & Chr$(5)
    cBuddyListMesg = Chr$(0) & Chr$(0) & Chr$(6) & Chr$(8)
    cYahooMail = Chr$(0) & Chr$(0) & Chr$(6) & Chr$(9)
    cYahooMesg = Chr$(0) & Chr$(0) & Chr$(7) & Chr$(0)
    cGrafiti = Chr$(0) & Chr$(0) & Chr$(8) & Chr$(0)

    sLogin = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(1)
    sLogout = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(2)
    sRoomEnter = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(17)
    sRoomLeave = Chr$(0) & Chr$(0) & Chr$(1) & Chr$(2)
    sInvitations = Chr$(0) & Chr$(0) & Chr$(1) & Chr$(7)
    sAwayBack = Chr$(0) & Chr$(0) & Chr$(2) & Chr$(1)
    sSpeak = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(65)
    sThink = Chr$(0) & Chr$(0) & Chr$(4) & Chr$(2)
    sEmote = Chr$(0) & Chr$(0) & Chr$(4) & Chr$(3)
    sAdvertisements = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(68)
    sPrivMesg = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(67)
    sBuddyListMesg = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(104)
    sYahooMail = Chr$(0) & Chr$(0) & Chr$(6) & Chr$(9)
    sYahooMesg = Chr$(0) & Chr$(0) & Chr$(7) & Chr$(0)
    SGrafiti = Chr$(0) & Chr$(0) & Chr$(8) & Chr$(0)


    AuthServer = "edit.my.yahoo.com"
    ChatServer = "cs.chat.yahoo.com"
    ChatPort = 8002
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wskChat.Close
    Unload frmAbout
    Unload frmHelp
End Sub

Private Sub lstCustomRooms_Click()
    txtStatus.Text = txtStatus.Text & vbCrLf & "Joining " & lstCustomRooms.List(lstCustomRooms.ListIndex)
    Dim theroom As String
    theroom = lstSubRooms.List(lstSubRooms.ListIndex)
    joinRoom theroom
End Sub

Private Sub lstMainCats_Click()
    If lstMainCats.ListIndex <> -1 Then
        getRooms vbNull, lstMainCats.ListIndex + 1
    End If
End Sub

Private Sub lstMainRooms_Click()
    If lstMainRooms.ListIndex <> -1 Then
        getRooms vbNull, lstMainCats.ListIndex + 1, subRoomsVariable(1, lstMainRooms.ListIndex)
    End If
End Sub

Private Sub lstSubRooms_Click()
    Dim theroom As String
    MsgBox "Joining " & lstSubRooms.List(lstSubRooms.ListIndex)
    theroom = lstSubRooms.List(lstSubRooms.ListIndex)
    joinRoom theroom
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAuthServer_Click()
    AuthServer = InputBox("Enter new authorization server:", "Auth Server Settings", AuthServer)
End Sub

Private Sub mnuChatServer_Click()
    ChatServer = InputBox("Enter a new chat server" & vbCrLf & "eg, cs1.yahoo.com, cs2.yahoo.com, cs3.yahoo.com, cs4.yahoo.com", "Chat Server Settings", ChatServer)
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show
End Sub

Private Sub mnuPort_Click()
    Dim vTemp As Variant
    Do
        vTemp = InputBox("Enter new chat port" & vbCrLf & "Default is 8002", "Chat Server Settings", ChatPort)
    Loop While Not IsNumeric(vTemp)
    ChatPort = CInt(vTemp)
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub txtPassword_Change()
    PassWord = txtPassword.Text
End Sub

Private Sub txtStatus_Change()
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub

Private Sub txtUsername_Change()
    UserName = txtUsername.Text
End Sub

Private Sub wskAuth_DataArrival(ByVal bytesTotal As Long)
    'ok, if theres an error, we don't care, drop the ball
    On Error Resume Next
    Static numTries As Integer
    If numTries = 5 Then
        MsgBox "ERROR: YahooChat, Could not get cookie, possibly invalid username/pass"
        numTries = 0
        If wskAuth.State <> 0 Then wskAuth.Close
        Exit Sub
    End If
    'theres no point in doin anythin if nothins there
    If Not bytesTotal = 0 Then
        'get that data, lets see if it fails
        wskAuth.GetData thePacket
        If wskAuth.State <> sckClosed Then wskAuth.Close
        If thePacket = "" Then 'it failed
            numTries = numTries + 1
            getCookie
            Exit Sub
        End If
        Dim startCookie As Integer
        Dim endCookie As Integer
        startCookie = InStr(1, thePacket, "Set-Cookie:", vbBinaryCompare) + Len("Set-Cookie: V=")
        Debug.Print thePacket
        If (startCookie = Len("Set-Cookie: V=")) Or _
           (InStr(1, thePacket, "ERROR", vbTextCompare) <> 0) Then 'if this happens then we didn't get a cookie
            thePacket = ""
            numTries = numTries + 1
            getCookie
            Exit Sub
        End If
        endCookie = InStr(startCookie, thePacket, ";", vbBinaryCompare)
        Cookie = Mid(thePacket, startCookie, endCookie - startCookie)
        gotCookie
    End If
End Sub

Private Sub getCookie()
    If wskAuth.State <> sckClosed Then wskAuth.Close
    wskAuth.RemoteHost = AuthServer
    wskAuth.RemotePort = 80
    Dim HTTPString As String
    HTTPString = "GET /config/ncclogin?.src=bl&" _
        & "login=" & UserName & "&passwd=" & PassWord _
        & "&n=1 HTTP/1.1" & Chr$(13) & Chr$(10) _
        & "Host: " & AuthServer & Chr$(13) & Chr$(10) _
        & "Accept: */*" & Chr$(13) & Chr$(10) _
        & Chr$(13) & Chr$(10)
    wskAuth.Connect AuthServer, 80
    Dim fatalErrorAK As Integer
    While wskAuth.State <> sckConnected
        fatalErrorAK = fatalErrorAK + 1
        DoEvents
        If fatalErrorAK = 30000 Then
            MsgBox "Could not connect to auth server"
            Exit Sub
        End If
    Wend
    wskAuth.SendData HTTPString
End Sub

Private Sub gotCookie()
    cmdAuth.Enabled = False
    cmdConnect.Enabled = True
    lblStatus.Caption = "Got Cookie (" & Cookie & ")"
    txtUsername.Enabled = False
    txtPassword.Enabled = False
End Sub

Private Function createPacket(Command As String, Content As String) As String
    Dim StringLen As String
    If Len(Content) > 65000 Then
        Err.Raise 10059, "YahooChat", "Content string is greater than 65000 chars"
        Exit Function
    ElseIf Len(Content) > 255 Then
        Dim tempInt1 As Integer
        Dim tempInt2 As Integer
        tempInt1 = Len(Content) Mod 256
        tempInt2 = (Len(Content) - tempInt1) / 256
        StringLen = Chr$(tempInt2) & Chr$(tempInt1)
    Else
        StringLen = Chr$(0) & Chr$(Len(Content))
    End If

    Dim createdPacket As String
    createdPacket = "YCHT" & cYahooVer _
        & Command & Chr$(0) & Chr$(0) & StringLen & _
        Content
    lastCommand = Command
    createPacket = createdPacket
End Function

Private Sub wskChat_DataArrival(ByVal bytesTotal As Long)
    'if theres an error, drop the ball
    If bytesTotal = 0 Then Exit Sub
    On Error Resume Next
    wskChat.GetData thePacket
    Dim sPacketType As String
    Dim sContent As String
    decodePacket thePacket, sPacketType, sContent
    Debug.Print thePacket
    Select Case sPacketType
        Case sLogin
            If InStr(1, thePacket, UserName, vbTextCompare) = 0 Then 'login failed
                MsgBox "couldn't login to server?"
                Exit Sub
            End If
            LoggedIn
        Case sRoomEnter
            If InStr(1, thePacket, UserName, vbTextCompare) = 0 Then
                MsgBox "Could not enter room"
                Exit Sub
            End If
            enteredRoom
        Case sSpeak
            txtStatus.Text = txtStatus.Text & vbCrLf & sContent
    End Select
End Sub

Private Sub joinRoom(sRoom As String)
    Dim roomPacket As String
    currentRoom = sRoom
    sRoom = "join" & Chr$(32) & sRoom
    roomPacket = createPacket(cRoomEnter, sRoom)
    wskChat.SendData roomPacket
End Sub

Private Sub LoggedIn()
    lblStatus.Caption = "Connected to " & wskChat.RemoteHost
    cmdJoinRoom.Enabled = True
End Sub

Private Sub enteredRoom()
    lblRoom.Caption = currentRoom
    txtSpeak.Enabled = True
    cmdSay.Enabled = True
    txtStatus.Text = txtStatus.Text & vbCrLf & "Joined " & currentRoom
End Sub

Private Sub Speak(toSay As String)
    'generate header
    Dim SpeakPacket As String
    SpeakPacket = createPacket(cSpeak, currentRoom & Chr$(1) & toSay)
    wskChat.SendData SpeakPacket
End Sub

'this function is intrinsically linked to wskRooms
Private Sub getRooms(saveRoomsToStringArray As String, _
  Optional Categories As Integer = 1, Optional SubCat As String)
  'CATEGORIES LISTING
  '1 = Adult Chat
  '2 = Computers & Internet
  '3 = Business & Finance
  '4 = Cultures & Community
  '5 = Countries and Cultures
  '6 = Teen
  '7 = Entertainment & Arts
  '8 = Family & Home
  '9 = Freinds
  '10 = Games
  '11 = Government & Politics
  '12 = Health & Wellness
  '13 = Hobbies & Crafts
  '14 = Music
  '15 = Recreation
  '16 = Regional
  '17 = Religon & Beliefs
  '18 = Romance
  '19 = Schools & Education
  '20 = Science
  
  Dim strCategories As String
  Select Case Categories
        Case 20
            Room = ".rmcat=1600082641"
        Case 19
            Room = ".rmcat=1600077623"
        Case 18
            Room = ".rmcat=1600083763"
        Case 17
            Room = ".rmcat=1600073831"
        Case 16
            Room = ".rmcat=1600043463"
        Case 15
            Room = ".rmcat=1600064623"
        Case 14
            Room = ".rmcat=1600417569"
        Case 13
            Room = ".rmcat=1600062280"
        Case 12
            Room = ".rmcat=1600060813"
        Case 11
            Room = ".rmcat=1600059353"
        Case 10
            Room = ".rmcat=1600052895"
        Case 9
            Room = ".rmcat=1600047754"
        Case 8
            Room = ".rmcat=1600038063"
        Case 7
            Room = ".rmcat=1600016068"
        Case 6
            Room = ".rmcat=1600008562"
        Case 5
            Room = ".rmcat=1600013556"
        Case 4
            Room = ".rmcat=1600008033"
        Case 3
            Room = ".rmcat=1600000002"
        Case 2
            Room = ".rmcat=1600004725"
        Case 1
            Room = ".rmcat=1600083764"
    End Select
    strCategories = Room
    If SubCat <> vbNullString Then
        strCategories = strCategories & "&exp=" & SubCat & "#" & SubCat
        subRoom = SubCat
    End If
    If strCategories = vbNullString Then
        Err.Raise 10059, "YahooChat", "Optional Category must be between 1 and 20"
        Exit Sub
    End If
    ReDim RoomsTempVariable(0 To 0)
    RoomsTempVariable(0) = strCategories
    
    'first lets make my job easier :)
    'Dim theArray() As String
    'Set theArray = saveRoomsToStringArray
    If wskRooms.State <> sckClosed Then wskRooms.Close
    wskRooms.Connect "chat.yahoo.com", 80
    Dim uhOh As Integer
    While wskRooms.State <> sckConnected
        DoEvents
        uhOh = uhOh + 1
        Sleep (10)
        If uhOh = 30000 Then
            Err.Raise 10059, "YahooChat", "could not connect to chat.yahoo.com"
            Exit Sub
        End If
    Wend
    Dim HTTPString As String
    HTTPString = "GET /c/roomlist/newlist.html?" & strCategories & " HTTP/1.1" & Chr$(13) & Chr$(10) _
        & "Host: chat.yahoo.com" & Chr$(13) & Chr$(10) _
        & "Accept: */*" & Chr$(13) & Chr$(10) _
        & Chr$(13) & Chr$(10)
    wskRooms.SendData HTTPString
End Sub

Private Sub wskRooms_DataArrival(ByVal bytesTotal As Long)
    Static theGrandTotal As String
    If bytesTotal <> 0 Then
        If roomsMode = False Then
            Dim theResponse As String
            On Error Resume Next
            wskRooms.GetData theResponse
            theGrandTotal = theGrandTotal & theResponse
            If InStr(1, theGrandTotal, "</html>", vbBinaryCompare) = 0 Then Exit Sub 'oh shit, theres more stuff on the way
            'else we made we it this far, so lets parse the html
            ReDim roomsVariable(0) As String
            ReDim subRoomsVariable(0 To 1, 0) As String
                
            Dim lenGrand As Integer
            lenGrand = Len(theGrandTotal)
            Dim startRoom As Long
            startRoom = 1
            Dim secondWhile As Integer
            While startRoom <> 0
                startRoom = InStr(startRoom, theGrandTotal, "name=", vbBinaryCompare)
                If startRoom <> 0 Then 'ok we are still getting the sub rooms
                    Dim RoomNumber As String
                    startRoom = startRoom + 1 + Len("name=")
                    RoomNumber = Mid$(theGrandTotal, startRoom, 10)
                    ReDim Preserve subRoomsVariable(0 To 1, 0 To (UBound(subRoomsVariable, 2) + secondWhile))
                    subRoomsVariable(1, UBound(subRoomsVariable, 2)) = RoomNumber
                    Dim endRoom As Long
                    Dim RoomName As String
                    lenGrand = Len(theGrandTotal)
                    startRoom = InStr(startRoom, theGrandTotal, "<a href=", vbBinaryCompare)
                    startRoom = InStr(startRoom, theGrandTotal, ">", vbBinaryCompare) + 1
                    endRoom = InStr(startRoom, theGrandTotal, "<", vbBinaryCompare)
                    RoomName = Mid$(theGrandTotal, startRoom, endRoom - startRoom)
                    subRoomsVariable(0, UBound(subRoomsVariable, 2)) = RoomName
                    If secondWhile = 0 Then secondWhile = 1
                End If
            Wend
            
            startRoom = 1
            Dim firstWhile As Integer
            While startRoom <> 0
                startRoom = InStr(startRoom, theGrandTotal, "&#183", vbBinaryCompare)
                If startRoom <> 0 Then
                    startRoom = InStr(startRoom, theGrandTotal, ">", vbBinaryCompare)
                    endRoom = InStr(startRoom, theGrandTotal, "<", vbBinaryCompare)
                    Dim sRoomName As String
                    sRoomName = Mid$(theGrandTotal, startRoom + 1, endRoom - startRoom - 1)
                    ReDim Preserve roomsVariable(0 To UBound(roomsVariable) + firstWhile)
                    roomsVariable(UBound(roomsVariable)) = sRoomName
                    If firstWhile = 0 Then firstWhile = 1
                End If
            Wend
            theGrandTotal = vbNullString
            wskRooms.Close
            gotRooms
        Else
        'we are gettin the custom rooms
            On Error Resume Next
            Dim audixResponse As String
            Static audixTotal As String
            wskRooms.GetData audixResponse
            audixTotal = audixTotal & audixResponse
            If InStr(1, audixTotal, "</html>", vbBinaryCompare) = 0 Then Exit Sub 'oh shit, theres more stuff on the way
            startRoom = 1
            Dim thirdWhile As Integer
            ReDim customRoomsVariable(0) As String
            While startRoom <> 0
                startRoom = InStr(startRoom, audixTotal, "&#183", vbBinaryCompare)
                If startRoom <> 0 Then
                    startRoom = InStr(startRoom, audixTotal, ">", vbBinaryCompare)
                    endRoom = InStr(startRoom, audixTotal, "<", vbBinaryCompare)
                    RoomName = Mid$(audixTotal, startRoom + 1, endRoom - startRoom - 1)
                    ReDim Preserve customRoomsVariable(0 To UBound(customRoomsVariable) + thirdWhile)
                    customRoomsVariable(UBound(customRoomsVariable)) = RoomName
                    If thirdWhile = 0 Then thirdWhile = 1
                End If
            Wend
        roomsMode = False
        wskRooms.Close
        audixTotal = vbNullString
        gotCustomRooms
        End If
    End If
    If wskRooms.State = sckClosing Then wskRooms.Close
End Sub

Private Sub decodePacket(WholePacket As String, PacketType As String, Content As String)
    PacketType = Mid$(WholePacket, 9, 4)
    Content = Right$(WholePacket, Len(WholePacket) - 16)
End Sub

Private Sub gotRooms()
    'populate the lists.
    lstMainRooms.Clear
    txtStatus.Text = txtStatus.Text & vbCrLf & "Got " & UBound(subRoomsVariable, 2) + 1 & " categories"
    txtStatus.Text = txtStatus.Text & vbCrLf & "Got " & UBound(roomsVariable) + 1 & " rooms"
    Dim iCounter As Integer
    For iCounter = 0 To UBound(subRoomsVariable, 2)
        lstMainRooms.AddItem subRoomsVariable(0, iCounter)
    Next
    lstSubRooms.Clear
    For iCounter = 0 To UBound(roomsVariable)
        lstSubRooms.AddItem roomsVariable(iCounter)
    Next
    getCustomRooms
End Sub

Private Sub getCustomRooms()
    'ok, we are getting custom rooms
    'cheezy, lets close our wskrooms socket
    On Error Resume Next
    wskRooms.Close
    Dim HTTPString As String
    HTTPString = "GET /c/roomlist/auidx.html?" & Room & " HTTP/1.1" & Chr$(13) & Chr$(10) _
        & "Host: chat.yahoo.com" & Chr$(13) & Chr$(10) _
        & "Accept: */*" & Chr$(13) & Chr$(10) _
        & Chr$(13) & Chr$(10)
    wskRooms.Connect "chat.yahoo.com", 80
    
    Dim tempCounter As Integer
    While wskRooms.State <> sckConnected
        tempCounter = tempCounter + 1
        DoEvents
        Sleep (5)
    Wend
    roomsMode = True
    wskRooms.SendData HTTPString
End Sub

Private Sub gotCustomRooms()
    lstCustomRooms.Clear
    Dim iCounter As Integer
    txtStatus.Text = txtStatus.Text & vbCrLf & "Got " & UBound(customRoomsVariable) + 1 & " custom rooms"
    For iCounter = 0 To UBound(customRoomsVariable)
        lstCustomRooms.AddItem customRoomsVariable(iCounter)
    Next
End Sub
