Attribute VB_Name = "MConstants"
Option Explicit

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
Private Const TAPIERR_NOREQUESTRECIPIENT = -2&
Private Const TAPIERR_REQUESTQUEUEFULL = -3&
Private Const TAPIERR_INVALDESTADDRESS = -4&

Public Sub Dial(Frm As Form, Num As String)
Dim buff As String
Dim nResult As Long
    nResult = tapiRequestMakeCall&(Trim$(Num), CStr(Frm.Caption), Frm.txtLastName & ", " & Frm.txtFirstName, "")
        If nResult <> 0 Then
            buff = "Error dialing number : "
            Select Case nResult
                Case TAPIERR_NOREQUESTRECIPIENT
                    buff = buff & "No Windows Telephony dialing application is running and none could be started."
                Case TAPIERR_REQUESTQUEUEFULL
                    buff = buff & "The queue of pending Windows Telephony dialing requests is full."
                Case TAPIERR_INVALDESTADDRESS
                    buff = buff & "The phone number is Not valid."
                Case Else
                    buff = buff & "Unknown error."
            End Select
        End If
End Sub


Public Sub DragForm(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Public Sub OpenContact(Name As String)

Dim Another As New frmDetails
Dim YearDiff As Integer
Dim x As Integer
  
'Check if form already exists
For x = 0 To Forms.Count - 1
'If so, Exit sub
    If Forms(x).Caption = "Contacts - " & Name Then Forms(x).SetFocus: Exit Sub
Next x
    
    
    With frmMain.RS
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !LastName & ", " & !FirstName = Name Then
                Exit Do
            Else
                .MoveNext
            End If
        Loop
        
Dim BDate As Date
On Error Resume Next

    Another.Visible = False
    Another.Caption = "Contacts - " & Name
    Another.txtFirstName = !FirstName
    Another.txtLastName = !LastName
    Another.txtAddress1 = !Address1
    Another.txtAddress2 = !Address2
    Another.txtAddress3 = !Address3
    Another.txtCity = !City
    Another.txtPostcode = UCase(!Postcode)
    Another.txtData1 = !Data1
    Another.txtData2 = !Data2
    Another.txtData3 = !Data3
    Another.txtData4 = !Data4
    Another.txtEmail = !Email
    Another.txtWebsite = !Website
    Another.txtNotes = !Notes
    Another.txtNotes.TabIndex = 0
    Another.Tag = Name
    Another.cmbCat.ListIndex = !cat
    Another.cmbData1.ListIndex = !Combo1
    Another.cmbData2.ListIndex = !Combo2
    Another.cmbData3.ListIndex = !Combo3
    Another.cmbData4.ListIndex = !Combo4

    Load Another
    Another.Visible = True

    Another.SSTab1.Tab = 0
    Another.Show
    End With
    Exit Sub
End Sub
