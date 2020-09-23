VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3465
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3465
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   60
      Top             =   30
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":3679
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   60
      TabIndex        =   0
      Top             =   3000
      Width           =   5835
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Load frmLogin
    frmLogin.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Load frmLogin
    frmLogin.Show
    Unload Me
End Sub
