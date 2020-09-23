VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book Login"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2010
      TabIndex        =   5
      Top             =   2565
      Width           =   1290
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3420
      TabIndex        =   4
      Top             =   2565
      Width           =   1290
   End
   Begin VB.ComboBox cmbUsers 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Top             =   1050
      Width           =   2970
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1515
      Width           =   2970
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This label shows the password for the selected user, the visible property should be set to false before use."
      Height          =   420
      Left            =   75
      TabIndex        =   7
      Top             =   1965
      Width           =   4680
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   2520
      Width           =   1665
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   1065
      Width           =   855
   End
   Begin VB.Image imgBorder 
      Height          =   900
      Left            =   -15
      Picture         =   "frmLogin.frx":0000
      Top             =   -30
      Width           =   4770
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB2 As Database
Public RS2 As Recordset
Dim CurrentUser As String

Private Sub cmbUsers_Click()
    CheckPass (cmbUsers.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdLogin_Click()
If txtPassword.Text = lblPass.Caption Then
    PassCorrect
Else
    MsgBox "Invalid username or password.", vbCritical, "Error"
End If
End Sub

Public Sub PassCorrect()
    Load frmMain
    frmMain.Show
    frmMain.sbMain.Panels(1).Text = "Welcome " & cmbUsers.Text
    frmMain.Caption = "Address Book XP - [" & cmbUsers.Text & "]"
    Unload frmLogin
End Sub

Private Sub Form_Load()
Set frmLogin.DB2 = OpenDatabase(App.Path & "\Data.abd")

    LoadCombo

End Sub

Public Sub LoadCombo()
Set frmLogin.RS2 = frmLogin.DB2.OpenRecordset("SELECT * FROM Users ORDER BY Username DESC")

    If RS2.EOF = False Then
        RS2.MoveFirst
        While RS2.EOF = False
            cmbUsers.AddItem RS2.Fields(1)
            RS2.MoveNext
        Wend
        cmbUsers.ListIndex = 0
    End If

End Sub

Public Sub CheckPass(Username As String)
Dim varPass As String

    With frmLogin.RS2
        If .RecordCount = 0 Then Exit Sub
            .MoveFirst
        Do While Not .EOF
            If !Username = Username Then
        Exit Do
        Else
            .MoveNext
        End If
    Loop
        
    lblPass.Caption = !Password

    End With

    Exit Sub

End Sub

