VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSendSms 
   Caption         =   "Sms Sender"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendSms.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   345
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "Hide modem properties"
      Top             =   7800
      Width           =   1095
   End
   Begin SmsSender.GSM GSM 
      Left            =   120
      Top             =   7800
      _extentx        =   873
      _extenty        =   450
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Done"
      Height          =   345
      Left            =   6120
      TabIndex        =   8
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4920
      TabIndex        =   7
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   7095
      Begin VB.TextBox txtMessage 
         Height          =   1575
         Left            =   1200
         MaxLength       =   320
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox txtRecipient 
         Height          =   345
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Message"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recipient"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GSM Modem / Phone"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.OptionButton optVoice 
         Caption         =   "Cell"
         Height          =   210
         Left            =   6240
         TabIndex        =   14
         ToolTipText     =   "The gadget is actually a cellular phone mostly used for voice."
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optData 
         Caption         =   "Data"
         Height          =   210
         Left            =   5520
         TabIndex        =   13
         ToolTipText     =   "The gadget is a data card that is mainly used for data calls"
         Top             =   360
         Width           =   735
      End
      Begin MSComctlLib.ListView lstReport 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox cboDevices 
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Device"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "lblMessage"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   7800
      Width           =   810
   End
End
Attribute VB_Name = "frmSendSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboDevices_Click()
    On Error Resume Next
    ' read the devices connected to the computer
    lblMessage.Caption = "Reading device properties..."
    GSM.DeviceProperties cboDevices.Text, lstReport
    optData.Value = False
    optVoice.Value = False
    lblMessage.Caption = ""
    Err.Clear
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    GSM.DisConnect
    Unload Me
    Err.Clear
End Sub

Private Sub cmdHide_Click()
    Select Case cmdHide.Caption
    Case "Hide"
        cmdHide.Caption = "Show"
        Me.Height = 4800
        Frame1.Height = 855
        Frame2.Top = 1080
        lstReport.Visible = False
        cmdHide.Top = 3840
        cmdSend.Top = 3840
        cmdCancel.Top = 3840
        lblMessage.Top = 3840
    Case Else
        Frame2.Top = 5040
        Me.Height = 8745
        Frame1.Height = 4935
        cmdHide.Caption = "Hide"
        cmdHide.Top = 7800
        cmdSend.Top = 7800
        cmdCancel.Top = 7800
        lblMessage.Top = 7800
        lstReport.Visible = True
    End Select
End Sub

Private Sub cmdSend_Click()
    On Error Resume Next
    'send the message to the selected recipients
    Dim mResult As String
    Dim mPort As String
    Dim mSpeed As String
    Dim lPos As Long
    Dim mRow() As String
    Dim msgResult As Long
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spNumbers() As String
    Dim sendResult As New Collection
    txtRecipient.Text = Trim$(txtRecipient.Text)
    txtMessage.Text = Trim$(txtMessage.Text)
    If cboDevices.ListIndex = -1 Then
        Call MsgBox("The gsm modem type has not been chosem, please select data or voice.", vbOKOnly + vbExclamation + vbApplicationModal, "Gsm Modem Type")
        Err.Clear
        Exit Sub
    End If
    If Len(txtRecipient.Text) = 0 Then
        Call MsgBox("The recipient(s) cannot be blank." & vbCr & "Please enter the receipient(s) cellular phone number or email address!", vbOKOnly + vbExclamation + vbApplicationModal, "Recipient(s) Error")
        Err.Clear
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lblMessage.Caption = "Reading port number..."
    mPort = 0
    mSpeed = 0
    ' find the port that the modem is connected to
    lPos = GSM.LstViewFindItem(lstReport, "AttachedTo")
    If lPos > 0 Then
        mRow = GSM.LstViewGetRow(lstReport, lPos)
        mPort = GSM.ExtractNumbers(mRow(2))
    End If
    lblMessage.Caption = "Reading speed..."
    ' find the maximum connection rate as set as settings
    '"460800,n,8,1"
    lPos = GSM.LstViewFindItem(lstReport, "maxbaudratetoserialport")
    If lPos > 0 Then
        mRow = GSM.LstViewGetRow(lstReport, lPos)
        mSpeed = mRow(2)
    End If
TryConnection:
    Screen.MousePointer = vbHourglass
    lblMessage.Caption = "Connecting to the device..."
    ' connect to the modem/phone
    If GSM.Connected = True Then GoTo SendMessage
    mResult = GSM.Connect(mPort, mSpeed)
    If optData.Value = True Then
        GSM.ModemType = DataCard
    ElseIf optVoice.Value = True Then
        GSM.ModemType = Cellphone
    End If
    Select Case mResult
    Case "OK"
        lblMessage.Caption = "Reading message format..."
        ' change the message format to text if it is not
        mResult = GSM.SMS_MessageFormat(ReadFormat)
        If mResult = "PDU" Then
            lblMessage.Caption = "Setting message format to text..."
            GSM.SMS_MessageFormat TextFormat
        End If
    Case Else
        Screen.MousePointer = vbDefault
        msgResult = MsgBox("A connection to the device could not be established!", vbRetryCancel + vbExclamation + vbApplicationModal, "Connection Error")
        If msgResult = vbCancel Then Exit Sub
        GoTo TryConnection
    End Select
SendMessage:
    txtRecipient.Text = Replace$(txtRecipient, ";", ",")
    spNumbers = Split(txtRecipient.Text, ",")
    rsTot = UBound(spNumbers)
    For rsCnt = 0 To rsTot
        lblMessage.Caption = "Sending message to " & spNumbers(rsCnt) & "..."
        mResult = GSM.SMS_Send(spNumbers(rsCnt), txtMessage.Text)
        Select Case mResult
        Case "OK"
            sendResult.Add "Message was sent successfully to " & spNumbers(rsCnt)
        Case Else
        End Select
        DoEvents
    Next
    lblMessage.Caption = ""
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    ' load devices connected to the computer to the combobox
    lblMessage.Caption = ""
    GSM.LoadDevices cboDevices
    GSM.LogFile = App.Path & "\sms sender.log"
    Err.Clear
End Sub

Private Sub GSM_Response(ByVal Result As String)
    'Debug.Print Result
End Sub

Private Sub txtMessage_Change()
    On Error Resume Next
    ' display how many characters are left and enable suitable buttons
    If Len(txtMessage.Text) = 0 Then
        cmdSend.Enabled = False
    Else
        cmdSend.Enabled = True
    End If
    Err.Clear
End Sub
