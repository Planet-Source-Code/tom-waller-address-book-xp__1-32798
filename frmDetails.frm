VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDetails 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7365
   Begin TabDlg.SSTab SSTab1 
      Height          =   4995
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   8811
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmDetails.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtFirstName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtLastName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbCat"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAddress1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAddress2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCity"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPostcode"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAddress3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtData1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtData2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtData3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtData4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtEmail"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtWebsite"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbData1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbData2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbData3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmbData4"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "frmDetails.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtNotes"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.ComboBox cmbData4 
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
         Left            =   3645
         TabIndex        =   14
         Text            =   "Other"
         Top             =   2610
         Width           =   1185
      End
      Begin VB.ComboBox cmbData3 
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
         Left            =   3645
         TabIndex        =   12
         Text            =   "Pager"
         Top             =   2250
         Width           =   1185
      End
      Begin VB.ComboBox cmbData2 
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
         Left            =   3645
         TabIndex        =   10
         Text            =   "Mobile"
         Top             =   1890
         Width           =   1185
      End
      Begin VB.ComboBox cmbData1 
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
         Left            =   3645
         TabIndex        =   8
         Text            =   "Phone"
         Top             =   1530
         Width           =   1185
      End
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4485
         Left            =   -74910
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   405
         Width           =   7110
      End
      Begin VB.TextBox txtWebsite 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3690
         MaxLength       =   167
         TabIndex        =   17
         Top             =   3945
         Width           =   3420
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3690
         MaxLength       =   167
         TabIndex        =   16
         Top             =   3330
         Width           =   3420
      End
      Begin VB.TextBox txtData4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4860
         TabIndex        =   15
         Top             =   2595
         Width           =   2250
      End
      Begin VB.TextBox txtData3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4860
         TabIndex        =   13
         Top             =   2235
         Width           =   2250
      End
      Begin VB.TextBox txtData2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4860
         TabIndex        =   11
         Top             =   1890
         Width           =   2250
      End
      Begin VB.TextBox txtData1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4860
         TabIndex        =   9
         Top             =   1530
         Width           =   2250
      End
      Begin VB.TextBox txtAddress3 
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
         Left            =   225
         TabIndex        =   4
         Top             =   2970
         Width           =   3240
      End
      Begin VB.TextBox txtPostcode 
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
         Left            =   2430
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3600
         Width           =   1050
      End
      Begin VB.TextBox txtCity 
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
         Left            =   225
         TabIndex        =   5
         Top             =   3600
         Width           =   2115
      End
      Begin VB.TextBox txtAddress2 
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
         Left            =   225
         TabIndex        =   3
         Top             =   2610
         Width           =   3240
      End
      Begin VB.TextBox txtAddress1 
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
         Left            =   225
         TabIndex        =   2
         Top             =   2250
         Width           =   3240
      End
      Begin VB.ComboBox cmbCat 
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
         ItemData        =   "frmDetails.frx":05C2
         Left            =   3645
         List            =   "frmDetails.frx":05C4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   765
         Width           =   3420
      End
      Begin VB.TextBox txtLastName 
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
         Left            =   225
         TabIndex        =   1
         Top             =   1530
         Width           =   3240
      End
      Begin VB.TextBox txtFirstName 
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
         Left            =   225
         TabIndex        =   0
         Top             =   780
         Width           =   3240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3690
         TabIndex        =   28
         Top             =   3735
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3675
         TabIndex        =   27
         Top             =   3090
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone(s):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3645
         TabIndex        =   26
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Postcode:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2430
         TabIndex        =   25
         Top             =   3375
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City / Town:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   3375
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   23
         Top             =   2025
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3645
         TabIndex        =   22
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   1275
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   540
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

Private Sub Form_Load()
    FillCat
    FillCombo
End Sub

Public Sub FillCombo()

    cmbData1.AddItem "Phone"
    cmbData1.AddItem "Mobile"
    cmbData1.AddItem "Fax"
    cmbData1.AddItem "Pager"
    cmbData1.AddItem "Ext."
    cmbData1.AddItem "Other"
    
    cmbData2.AddItem "Phone"
    cmbData2.AddItem "Mobile"
    cmbData2.AddItem "Fax"
    cmbData2.AddItem "Pager"
    cmbData2.AddItem "Ext."
    cmbData2.AddItem "Other"
    
    cmbData3.AddItem "Phone"
    cmbData3.AddItem "Mobile"
    cmbData3.AddItem "Fax"
    cmbData3.AddItem "Pager"
    cmbData3.AddItem "Ext."
    cmbData3.AddItem "Other"
    
    cmbData4.AddItem "Phone"
    cmbData4.AddItem "Mobile"
    cmbData4.AddItem "Fax"
    cmbData4.AddItem "Pager"
    cmbData4.AddItem "Ext."
    cmbData4.AddItem "Other"

End Sub

Sub FillCat()
    cmbCat.AddItem "Friend"
    cmbCat.AddItem "Family"
    cmbCat.AddItem "Co-Worker"
    cmbCat.AddItem "General"
End Sub

Sub UpdateMe()
    With frmMain.RS
        .MoveFirst
            Do While Not .EOF
                If !LastName & ", " & !FirstName = Tag Then
                    Exit Do
                Else
                    .MoveNext
                End If
            Loop
    On Error Resume Next
        .Edit
            !FirstName = txtFirstName.Text
            !LastName = txtLastName.Text
            !Address1 = txtAddress1.Text
            !Address2 = txtAddress2.Text
            !Address3 = txtAddress3.Text
            !City = txtCity.Text
            !Postcode = txtPostcode.Text
            !Email = txtEmail.Text
            !Website = txtWebsite.Text
            !Combo1 = cmbData1.ListIndex
            !Combo2 = cmbData2.ListIndex
            !Combo3 = cmbData3.ListIndex
            !Combo4 = cmbData4.ListIndex
            !Data1 = txtData1.Text
            !Data2 = txtData2.Text
            !Data3 = txtData3.Text
            !Data4 = txtData4.Text
            !Notes = txtNotes.Text
            !cat = cmbCat.ListIndex
        .Update
    End With

    frmContacts.LoadContacts

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Answer As Integer

Answer = MsgBox("Do you want to update this record?", vbYesNoCancel + vbQuestion, "Update")
    If Answer = vbYes Then
        UpdateMe
        Unload Me
    ElseIf Answer = vbCancel Then
        Cancel = True
    End If
    

End Sub

