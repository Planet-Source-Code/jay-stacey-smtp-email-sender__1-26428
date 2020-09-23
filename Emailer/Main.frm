VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form mainfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SendEmail"
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5805
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Logtxt 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   5775
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5027
            MinWidth        =   5027
            Text            =   "Hackragent - Email Sender"
            TextSave        =   "Hackragent - Email Sender"
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/08/2001"
            Key             =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "2:03 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu settings 
         Caption         =   "Settings"
      End
      Begin VB.Menu SendEmail 
         Caption         =   "Send Email"
      End
      Begin VB.Menu refreshstatus 
         Caption         =   "Refresh Status"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exit_Click()
End ' End
End Sub

Private Sub Form_Load()
Logtxt.Text = Logtxt.Text & "Welcome : " & WinSock.LocalIP & " , " & WinSock.LocalHostName & Chr$(13) & Chr$(10) 'Display Welcome message
Logtxt.Text = Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10) ' Seperator
End Sub

Private Sub Logtxt_Change()
Logtxt.SelStart = Len(Logtxt.Text) ' keep view on the last character of the textbox
End Sub

Private Sub refreshstatus_Click()
Logtxt.Text = "" ' Refresh Textbox
End Sub

Private Sub SendEmail_Click()
SendMail ' Call SendEmail Sub
End Sub

Private Sub settings_Click()
EmailSettings.Show ' Show the EmailSettings Form
mainfrm.Enabled = False ' MainFrm cannot be changed while Emailsettings form is in view
End Sub

Private Sub SendMail()

If COPYS <= 1 Then ' If copys = less than 1 then log and exit sub
Logtxt.Text = Logtxt.Text & "Invalid Number of copys.." & Chr$(13) & Chr$(10)
Logtxt.Text = Logtxt.Text & "Aborted on " & Date & " at " & Time & Chr$(13) & Chr$(10) 'Port Closure Success
Logtxt.Text = Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10) 'Port Closure Success
Exit Sub ' exit sub
Else

For i = 1 To COPYS ' repeat till i = copys

Logtxt.Text = Logtxt.Text & "Checking settings..." & Chr$(13) & Chr$(10)

If SMTP_HOST = "" Or MAIL_TO = "" Then 'If SMTP_HOST or MAIL_TO = "" then log and exit sub
Logtxt.Text = Logtxt.Text & "Settings check Failed" & Chr$(13) & Chr$(10)
Logtxt.Text = Logtxt.Text & "Reconfigure Email Settings.." & Chr$(13) & Chr$(10)
Logtxt.Text = Logtxt.Text & "Aborted on " & Date & " at " & Time & Chr$(13) & Chr$(10)
Logtxt.Text = Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10) 'Port Closure Success
StatusBar.Panels(1).Text = "Status : " & "Error! Aborted..." ' show abort in statusbar
Exit Sub ' exit sub
Else
Logtxt.Text = Logtxt.Text & "Settings check passed!" & Chr$(13) & Chr$(10) ' SMTP check passed
End If

StatusBar.Panels(1).Text = "Status : " & "Sending mail - " & i & " of " & COPYS ' Show the number of emails sent so far

WinSock.Close 'Close Winsock Port

WinSock.Connect SMTP_HOST, PORT 'Connect to SMTP_HOST on PORT
Logtxt.Text = Logtxt.Text & "Connecting to : " & SMTP_HOST & " on port : " & PORT & Chr$(13) & Chr$(10)  'Display Current Status

Do While WinSock.State <> sckConnected 'loop until connection established
DoEvents
If WinSock.State = sckError Then ' if winsock has an error log and exit sub
Logtxt.Text = Logtxt.Text & "Error Connecting to :" & SMTP_HOST & " on port : " & PORT & Chr$(13) & Chr$(10)
WinSock.Close
Logtxt.Text = Logtxt.Text & "Aborted Successfully on " & Date & " at " & Time & Chr$(13) & Chr$(10) 'Port Closure Success
Logtxt.Text = Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10) 'Port Closure Success
StatusBar.Panels(1).Text = "Status : " & "Error! Aborted..."
Exit Sub
Else
End If
Loop

Logtxt.Text = Logtxt.Text & "Connected to : " & SMTP_HOST & " on " & Date & " at " & Time & Chr$(13) & Chr$(10) ' Connection established

Logtxt.Text = Logtxt.Text & "Waiting for reply..." & Chr$(13) & Chr$(10)

Do While Green_Light = False ' wait for incomming data before proceding
DoEvents
Loop

Green_Light = False ' set greenlight to false

'-------------------------------------------------------
WinSock.SendData "HELO " & WinSock.RemoteHostIP & Chr$(13) & Chr$(10) 'Send Helo command to SMTP_HOST
Logtxt.Text = Logtxt.Text & "Hello sent.. " & Chr$(13) & Chr$(10) ' Alert user
Do While Green_Light = False
DoEvents
Loop

Green_Light = False
'--------------------------------------------------------

WinSock.SendData "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10) 'Send MAIL_FROM to SMTP_HOST

Logtxt.Text = Logtxt.Text & "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10)

Do While Green_Light = False 'Wait for DataArrival
DoEvents
Loop

Green_Light = False

WinSock.SendData "RCPT TO: " & MAIL_TO & Chr$(13) & Chr$(10) 'Send MAIL_TO to SMTP_HOST

Logtxt.Text = Logtxt.Text & "MAIL TO: " & MAIL_TO & Chr$(13) & Chr$(10)

Do While Green_Light = False 'Wait for DataArrival
DoEvents
Loop

Green_Light = False

WinSock.SendData "DATA" & Chr$(13) & Chr$(10) 'Send DATA to SMTP_HOST

Logtxt.Text = Logtxt.Text & "DATA" & Chr$(13) & Chr$(10)

Do While Green_Light = False 'Wait for DataArrival
DoEvents
Loop

Green_Light = False

WinSock.SendData "FROM: " & SENDER_NAME & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10) ' Send senders name
WinSock.SendData "TO: " & RECIEVER_NAME & " <" & MAIL_TO & ">" & Chr$(13) & Chr$(10) ' send recievers name
WinSock.SendData "SUBJECT: " & SUBJECT & Chr$(13) & Chr$(10) 'send subject
WinSock.SendData Chr$(13) & Chr$(10)
WinSock.SendData DATA & Chr$(13) & Chr$(10) ' Send DATA to SMTP_HOST

Logtxt.Text = Logtxt.Text & DATA & Chr$(13) & Chr$(10)

WinSock.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)

Logtxt.Text = Logtxt.Text & Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)

Do While Green_Light = False 'Wait for DataArrival
DoEvents
Loop

Green_Light = False

WinSock.SendData "QUIT" & Chr$(13) & Chr$(10) ' Send Quit command
Logtxt.Text = Logtxt.Text & "Email Sent Successfully..." & Chr$(13) & Chr$(10) 'Success
WinSock.Close
Logtxt.Text = Logtxt.Text & "Completed Successfully on " & Date & " at " & Time & Chr$(13) & Chr$(10)
Logtxt.Text = Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10)
Next i ' Repeat process
StatusBar.Panels(1).Text = "Status : " & COPYS & " Emails Sent Successfully.." ' Show how many emails have sent successfully

End If

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

WinSock.GetData DATAFile 'Recieve Reply "DATAFile" from SMTP_HOST
Reply = Mid(DATAFile, 1, 3)
Logtxt.Text = Logtxt.Text & DATAFile & Chr$(13) & Chr$(10)


If Reply = 220 Or Reply = 250 Or Reply = 354 Then 'if reply is OK then greenlight = true
Green_Light = True
Logtxt.Text = Logtxt.Text & "DataArrival Accepted.." & Chr$(13) & Chr$(10)
Else ' Else if reply is other than OK log and close winsock
Logtxt.Text = Logtxt.Text & "DataArrival Denied.." & DATAFile & Chr$(13) & Chr$(10)
WinSock.Close ' close port
Logtxt.Text = Logtxt.Text & "Aborted Successfully on " & Date & " at " & Time & Chr$(13) & Chr$(10)
Logtxt.Text = Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10)
StatusBar.Panels(1).Text = "Status : " & "Error! Aborted..."
End If

End Sub

Private Sub WinSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Logtxt.Text = Logtxt.Text & "Winsock Error Description : " & Description & " on " & Date & " at " & Time & Chr$(13) & Chr$(10) 'Log Winsock Error Description
Logtxt.Text = Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10)
StatusBar.Panels(1).Text = "Status : " & "Error! Aborted..."
End Sub
