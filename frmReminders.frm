VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReminders 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminders"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmReminders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4350
   Begin MSComctlLib.StatusBar sbReminders 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   4620
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgReminders 
      Left            =   3315
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReminders.frx":014A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbReminders 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   635
      ButtonWidth     =   2487
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgReminders"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Reminder"
            Object.ToolTipText     =   "Add a reminder"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstReminders 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4305
      ItemData        =   "frmReminders.frx":02A4
      Left            =   -15
      List            =   "frmReminders.frx":02AB
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   345
      Width           =   4395
   End
End
Attribute VB_Name = "frmReminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Table As Recordset
Dim Active As Boolean

Private Sub Form_Load()
    LoadItems
End Sub

Private Sub LoadItems()
Active = False
Dim Count As Integer

    Count = 0
    lstReminders.Clear

Set Table = frmMain.DB.OpenRecordset("SELECT * FROM Reminders ORDER BY ITEM DESC")

    With Table
        If .RecordCount = 0 Then Exit Sub
            .MoveFirst
        Do While Not .EOF
            lstReminders.AddItem !item, Count
            lstReminders.Selected(Count) = !Done
            .MoveNext
            Count = Count + 1
        Loop
    End With
    
    Active = True
    sbReminders.SimpleText = lstReminders.ListCount & " items"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    DragForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Public Sub AddReminder()
On Error GoTo err:

Dim Answer As String
Dim DateFor As String
    
    Answer = InputBox("Enter To Do Item:", "New Item")
    DateFor = InputBox("Enter Due Date: ( DD/MM/YYYY )", "Date Due")
  
    If DateFor = "" Then Exit Sub
    If Answer = "" Then Exit Sub
Set Table = frmMain.DB.OpenRecordset("SELECT * FROM Reminders ORDER BY DATE DESC")

    With Table
        .AddNew
            !item = Answer
            !DateFor = DateFor
            !Done = False
            !Date = Date
        .Update
    End With
    
    LoadItems

err:
MsgBox "Please use the correct formatting.", vbCritical, "Error"
End Sub

Private Sub lstReminders_ItemCheck(item As Integer)
    If Active = False Then Exit Sub
Dim Checked As Integer

    If lstReminders.Selected(item) Then Checked = True
    If Checked Then
        CheckItem item
    Else
        UnCheckItem item
    End If
End Sub


Sub CheckItem(item As Integer)
Set Table = frmMain.DB.OpenRecordset("SELECT * FROM Reminders ORDER BY ITEM DESC")
    
    With Table
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !item = lstReminders.List(item) Then
                .Edit
                    !Done = True
                .Update
            Exit Do
        Else
            .MoveNext
        End If
        Loop

Dim Asnwer As Integer
    Answer = MsgBox("Delete From List?" & vbCrLf & lstReminders.List(item), vbYesNo + vbQuestion, "Item Done")
    If Answer = vbNo Then Exit Sub
    .Delete
    LoadItems
    End With
End Sub

Sub UnCheckItem(item As Integer)
    With Table
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !item = lstReminders.List(item) Then
                .Edit
                    !Done = False
                .Update
        Exit Do
        Else
            .MoveNext
        End If
        Loop
    LoadItems
    End With
End Sub

Private Sub tbReminders_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            AddReminder
        Case Else
            Exit Sub
    End Select
End Sub
