VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Address Book XP"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   420
      Left            =   4515
      TabIndex        =   2
      Top             =   5490
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAbout.frx":0000
      Top             =   3315
      Width           =   4425
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   315
      X2              =   5805
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   315
      X2              =   5805
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled: Microsoft VB Ent. 6.0 - 18/03/2002"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   270
      TabIndex        =   5
      Top             =   5055
      Width           =   5475
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author: Tom Waller"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   270
      TabIndex        =   4
      Top             =   4785
      Width           =   5475
   End
   Begin VB.Label lblVer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   3
      Top             =   4500
      Width           =   5475
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   5760
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   270
      X2              =   5760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Image Image1 
      Height          =   2970
      Left            =   -75
      Picture         =   "frmAbout.frx":01B9
      Top             =   -15
      Width           =   6075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Optimized for the AMD Athlon XP Proccessor and Windows XP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   45
      TabIndex        =   0
      Top             =   5490
      Width           =   3450
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVer.Caption = App.Title & " Version " & App.Major & "  Release " & App.Minor
End Sub

