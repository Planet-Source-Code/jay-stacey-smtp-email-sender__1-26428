VERSION 5.00
Begin VB.Form EmailSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox COPYStxt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Text            =   "1"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton CANCELbutton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton OKbutton 
      Caption         =   "Ok"
      Height          =   375
      Left            =   9000
      TabIndex        =   19
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Email"
      Height          =   3495
      Left            =   5040
      TabIndex        =   14
      Top             =   120
      Width           =   5295
      Begin VB.TextBox DATAtxt 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox SUBJECTtxt 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label8 
         Caption         =   "Email Base :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Subject :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Recievers Details"
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   4815
      Begin VB.TextBox RECIEVERNAMEtxt 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox MAILTOtxt 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Recievers Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Recievers Email :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Senders Details"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
      Begin VB.TextBox MAILFROMtxt 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox SENDERNAMEtxt 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Senders Email :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Senders Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Smtp Settings"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox SMTPtxt 
         Height          =   315
         ItemData        =   "EmailSettings.frx":0000
         Left            =   1200
         List            =   "EmailSettings.frx":002E
         TabIndex        =   23
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox PORTtxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "25"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Port :"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Smtp Server :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Number of copys :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "EmailSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CANCELbutton_Click()
mainfrm.Show ' show MainFrm
mainfrm.Enabled = True
EmailSettings.Hide '
End Sub

Private Sub Form_Load()
'Set variables
PORT = 25
SMTPtxt.Text = SMTP_HOST
PORTtxt.Text = PORT
MAILTOtxt.Text = MAIL_TO
MAILFROMtxt.Text = MAIL_FROM
SUBJECTtxt.Text = SUBJECT
SENDERNAMEtxt.Text = SENDER_NAME
RECIEVERNAMEtxt.Text = RECIEVER_NAME
DATAtxt.Text = DATA

End Sub

Private Sub OKbutton_Click()
mainfrm.Show
mainfrm.Enabled = True

'Set variables
SMTP_HOST = SMTPtxt.Text
PORT = PORTtxt.Text
MAIL_TO = MAILTOtxt.Text
MAIL_FROM = MAILFROMtxt.Text
SUBJECT = SUBJECTtxt.Text
SENDER_NAME = SENDERNAMEtxt.Text
RECIEVER_NAME = RECIEVERNAMEtxt.Text
DATA = DATAtxt.Text
COPYS = COPYStxt.Text



mainfrm.Logtxt.Text = mainfrm.Logtxt.Text & "Settings Updated Successfully on " & Date & " at " & Time & Chr$(13) & Chr$(10) 'Port Closure Success
mainfrm.Logtxt.Text = mainfrm.Logtxt.Text & "-----------------------------------------------------------------------------" & Chr$(13) & Chr$(10)
mainfrm.WinSock.Close
mainfrm.StatusBar.Panels(1).Text = "Status : " & "Updated Settings Successfull.."
EmailSettings.Hide
End Sub

Private Sub PORTtxt_DblClick()
PORTtxt.Locked = False
End Sub
