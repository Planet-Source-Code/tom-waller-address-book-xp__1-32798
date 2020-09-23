VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4365
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10716
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1005
            MinWidth        =   1005
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuShowContactList 
         Caption         =   "Show / Hide Contact &List"
      End
      Begin VB.Menu mnuToDo 
         Caption         =   "Show / Hide &To Do List"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   [ R E A D M E ]

'   Thank you for choosing Address Book XP. This program has been designed
'   to manage your phone book and diary appointments. It has been designed
'   with the UK in mind as that is my nationality, but it may be altered to suit
'   your own preferences.

'   Before I continue, credits must be issued to SparQ, whose code inspired me
'   and influenced the style of this program. All the code has been re-written for
'   Windows XP but some similarities may be spotted.

'   If you are using this program without the system fonts i have used then you
'   will experience some problems with screen layouts. If you are using WinXP
'   Then you should not experience any problems, otherwise, install the fonts
'   included in the zip file.

'   Once again, thank you for downloading my code, I hope you have loads
'   of good ideas on how to improve it. If you do improve it, please make some
'   acknowlegement to me, as I have done to SparQ above. Please leave some
'   comments on my submission.

'   [ Tom Waller ]
'   [ tom8572@hotmail.com]
'   [ 18/03/2002 ]
'   [ Visual Basic 6.0 ENT. ]

Public DB As Database       'This is the variable name that represens the database
Public RS As Recordset       'and the same for the recordset to be modified

Private Sub MDIForm_Load()
   
    mnuShowContactList_Click
    mnuToDo_Click
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Public Sub mnuShowContactList_Click()
    If Not frmContacts.Visible Then
        Load frmContacts
        frmContacts.Top = 60
        frmContacts.Left = 60
        frmContacts.Show
    Else
        Unload frmContacts
    End If
End Sub

Private Sub mnuToDo_Click()
    If Not frmReminders.Visible Then
        Load frmReminders
        frmReminders.Top = 60
        frmReminders.Left = Screen.Width - (frmReminders.Width + 160)
        frmReminders.Show
    Else
        Unload frmReminders
    End If
End Sub
