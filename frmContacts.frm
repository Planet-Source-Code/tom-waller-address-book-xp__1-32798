VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContacts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   3120
   Visible         =   0   'False
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   375
      TabIndex        =   3
      ToolTipText     =   "Search for a contact"
      Top             =   375
      Width           =   2730
   End
   Begin MSComctlLib.StatusBar sbContacts 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   4545
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imgContacts 
      Left            =   945
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":06E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":0C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":1018
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":13B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbContacts 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   582
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgContacts"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            Object.ToolTipText     =   "Add a new contact"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete   "
            Object.ToolTipText     =   "Delete a contact"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   -15
      TabIndex        =   0
      Top             =   705
      Width           =   3165
   End
   Begin VB.Label lblViewType 
      Height          =   270
      Left            =   930
      TabIndex        =   4
      Top             =   390
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   60
      Picture         =   "frmContacts.frx":174C
      Top             =   390
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   0
      Picture         =   "frmContacts.frx":1AD6
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For selecting a contact from the contact list:
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const LB_FINDSTRING = &H18F
' For setting up a thin border on a picture box control:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

'This function creates a flat looking control. I have used it on the Search text
'box at the top of frmContacts. In order for this to work, your controls
'apperence property must be set to flat with no border if possible otherwise
'you will experience some weird looking controls!

Private Function ThinBorder(ByVal lhWnd As Long, ByVal bState As Boolean)
Dim lS As Long

   lS = GetWindowLong(lhWnd, GWL_EXSTYLE)
   If Not (bState) Then
      lS = lS Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
   Else
      lS = lS Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
   End If
   SetWindowLong lhWnd, GWL_EXSTYLE, lS
   SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Function

Private Sub Form_Load()
    
    'Select the data source to be used. In this case it's a Access Database
    'renamed from .mdb to .abd
    
    Set frmMain.DB = OpenDatabase(App.Path & "\data.abd")
    Set frmMain.RS = frmMain.DB.OpenRecordset("SELECT * FROM CONTACTS ORDER BY LastName DESC")
    
    ThinBorder txtSearch.hwnd, True

    frmContacts.Height = 5227
    frmContacts.Width = 3225
    
    LoadContacts
    
End Sub

Sub LoadContacts()
    List1.Clear
    
    With frmMain.RS
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
            Do While Not .EOF
                List1.AddItem !LastName & ", " & !FirstName
                .MoveNext
            Loop
    End With

    sbContacts.SimpleText = List1.ListCount & " contacts"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    DragForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.mnuShowContactList_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then DragForm Me
End Sub

Private Sub Label3_Click()
    With frmMain.RS
        .AddNew
            !FirstName = "Contact"
            !LastName = "New"
        .Update
    End With
    
    LoadContacts
    OpenContact "New, Contact"
End Sub

Private Sub List1_DblClick()
    OpenContact List1.Text
End Sub

Private Sub tbContacts_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Answer As Integer

    Select Case Button.Index
        Case 1
            With frmMain.RS
                .AddNew
                   !FirstName = "Contact"
                    !LastName = "New"
                .Update
            End With

    LoadContacts
    OpenContact "New, Contact"

        Case 2
            If List1.ListIndex < 0 Then Exit Sub
            On Error GoTo err
            Answer = MsgBox("Delete: " & List1.Text & "?", vbYesNo + vbQuestion, "Delete")
                If Answer = vbYes Then
                    With frmMain.RS
                        .MoveFirst
                        Do Until !LastName & ", " & !FirstName = List1.Text
                             .MoveNext
                        Loop
                            .Delete
            End With
            End If

LoadContacts

Exit Sub
err:
MsgBox "Error Deleting: " & List1.Text, vbCritical, "Error"
End Select
End Sub

'This is the function for picking records out of the list box at run-time

Private Sub txtSearch_Change()
    List1.Visible = True
    List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, _
    ByVal txtSearch.Text)
End Sub
