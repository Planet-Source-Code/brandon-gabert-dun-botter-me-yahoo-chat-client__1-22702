VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "http://chat.yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim theMsg As String
    theMsg = "Ok, you came here for help and here it is, " _
        & vbCrLf & "1. Enter your username and password for " _
        & "yahoochat, you can get these by going to chat.yahoo.com " _
        & "clicking 'sign-up now', entering some info.  " _
        & "you can figure it out I am sure. " _
        & "even try their java client, it works pretty well, cheetachat " _
        & "is a step above though for sure" & vbCrLf _
        & "2. Click authorize in my program, this will " _
        & "retrieve a cookie from yahoo used to authenticate " _
        & "yourself when you login" & vbCrLf _
        & "3. Click connect, this will officially connect you to " _
        & "the chat server, good going if you made it this far. " _
        & "4. Click on one of the main categories, this will fill " _
        & "the other listboxes, click on a main rooms listbox, this will " _
        & "refill the other two listboxes, click on one of those two " _
        & "and you will join a room, alternatively, click join room " _
        & "type it into the inputbox... but I included the rooms in this " _
        & "release. parsing html, fun fun :-)" & vbCrLf _
        & "5. Chat away using the speak box, not the friendliest app " _
        & "to chat with, by any means, but its a proof of concept project " _
        & "really.  Wanted to see if I could do it.  First winsock project " _
        & "completed about three months ago, but never got to share it with the " _
        & "world for various reasons, officially released two weeks ago " _
        & "on my site pumkinhed.com, but I thought I would share with you " _
        & "planet-source dudes.  Thanks for the exposure, this app may be a "
    theMsg = theMsg & "bit buggy (watch for the danger will robinson packet, still not " _
        & "100% sure whats going on there, oh well i guess, thats the end of help."
    Label1.Caption = theMsg
End Sub

Private Sub Label2_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "http://chat.yahoo.com"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL
End Sub
