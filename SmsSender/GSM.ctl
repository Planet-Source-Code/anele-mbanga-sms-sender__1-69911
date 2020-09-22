VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl GSM 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ControlContainer=   -1  'True
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GSM.ctx":0000
   ScaleHeight     =   510
   ScaleWidth      =   495
   ToolboxBitmap   =   "GSM.ctx":0442
   Begin MSCommLib.MSComm MSComm 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "GSM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mSpeed As String
Event Response(ByVal Result As String)
Private mCommPort As Integer
Private mPortOpen As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private m_Buffer As String
Private mLogFile As String
Private mReadDelete As String
Private mWriteSend As String
Private mReceived As String
Private mReceivedUsed As Long
Private mReceivedCapacity As Long
Private mReadDeleteUsed As Long
Private mReadDeleteCapacity As Long
Private mWriteSendUsed As String
Private mWriteSendCapacity As String
Private mpbMemory As String
Private mpbCapacity As Long
Private mpbUsed As Long
Private resultPos As Long
'>>>>>>>>>>> obtained from planet source code to receive
'>>>>>>>>>>> pdu messages
Private mvarnoSCA As String
Private mvarnoOA As String
Private mvarFO As String
Private mvarDCS As String
Private mvarSCTS_Tgl As String
Private mvarSCTS_Jam As String
Private mvarSCTS_Tgl_A As String
Private mvarSCTS_Jam_A As String
Private mvarIndexSend As String
Private mvarUDL As String
'>>>>>>>>>>>>>
Public Enum FindWhere
    Search_Text = 0
    Search_SubItem = 1
    Search_Tag = 2
End Enum
Public Enum SearchType
    Search_Partial = 1
    Search_Whole = 0
End Enum
Public Enum SmsMemoryStorageEnum
    SimMemory = 0
    MobileEquipmentMemory = 1
    BothMemories = 2
    ReadMemorySetting = 3
    BroadcastMessageStorage = 4
    StatusReportStorage = 5
    TerminalAdapterStorage = 6
End Enum
Public Enum MessageFormatEnum
    TextFormat = 1
    PDUFormat = 0
    ReadFormat = 2
End Enum
Public Enum EchoEnum
    EchoOff = 0
    EchoOn = 1
End Enum
Public Enum PhoneBookMemoryStorageEnum
    SimPhoneBook = 0
    MobileEquipmentPhoneBook = 1
    BothPhoneBooks = 2
    ReadPhoneBookSetting = 3
    SupportedMemory = 4
End Enum
Public Enum CentreNumberEnum
    SetCentreNumber = 0
    ReadCentreNumber = 1
End Enum
Public Enum SmsTypesEnum
    RecRead = 0
    RecUnread = 1
    Rec = 2
    StoSent = 3
    StoUnsent = 4
    All = 5
End Enum
Public Enum PDUFlashMessageTypeEnum
    ' flash = 00, non flash = F0
    NonFlash = 0
    Flash = 1
End Enum
Public Enum PDUDisplayReportSMSEnum
    ' yes = 31, no = 11
    YesReport = 1
    NoReport = 0
End Enum
Public Enum PDULimitPeriodOfDeliveryEnum
    ' 1=1hour,2=12hour,3=1day(default),4=2day,5=1week
    OneHour = 1
    TwelveHours = 2
    OneDay = 3
    TwoDays = 4
    OneWeek = 5
End Enum
Public Enum PhoneBookMemoryToFormatEnum
    SimPhoneBookFormat = 0
    MobileEquipmentPhoneBookFormat = 1
End Enum
Public Enum WriteLocationEnum
    WriteRecRead = 0
    WriteRecUnread = 1
    WriteStoSent = 3
    WriteStoUnsent = 4
End Enum
Public Enum ModemTypeEnum
    DataCard = 0
    Cellphone = 1
End Enum
Private nModemType As ModemTypeEnum
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Property Get CommPort() As Integer
    On Error Resume Next
    ' get the port used by the modem
    CommPort = mCommPort
    Err.Clear
End Property
Public Property Let CommPort(nCommPort As Integer)
    On Error Resume Next
    ' set the port to use for the modem connection
    MSComm.CommPort = nCommPort
    mCommPort = nCommPort
    PropertyChanged "CommPort"
    Err.Clear
End Property
Public Property Get ModemType() As ModemTypeEnum
    On Error Resume Next
    ' get the modem type
    ModemType = nModemType
    Err.Clear
End Property
Public Property Let ModemType(newModemType As ModemTypeEnum)
    On Error Resume Next
    ' set the modem type
    nModemType = newModemType
    If nModemType = Cellphone Then
        resultPos = 3
    ElseIf nModemType = DataCard Then
        resultPos = 2
    End If
    PropertyChanged "ModemType"
    Err.Clear
End Property
Private Function FixSmsDate(sDate As String) As String
    On Error Resume Next
    Dim syyyy As String
    Dim smm As String
    Dim sdd As String
    syyyy = MvField(sDate, 1, "/")
    smm = MvField(sDate, 2, "/")
    sdd = MvField(sDate, 3, "/")
    sdd = sdd & "/" & smm & "/" & syyyy
    FixSmsDate = Format$(sdd, "dd/mm/yyyy")
    Err.Clear
End Function
Public Property Get LogFile() As String
    On Error Resume Next
    ' get the path of the log file
    LogFile = mLogFile
    Err.Clear
End Property
Public Property Let LogFile(nLogFile As String)
    On Error Resume Next
    ' set the path of the log file
    mLogFile = nLogFile
    PropertyChanged "LogFile"
    Err.Clear
End Property
Public Property Get Settings() As String
    On Error Resume Next
    ' get the settings of the modem
    Settings = MSComm.Settings
    Err.Clear
End Property
Public Property Get VM() As String
    On Error Resume Next
    ' value delimiter
    VM = Chr$(253)
    Err.Clear
End Property
Public Property Get FM() As String
    On Error Resume Next
    ' value delimiter
    FM = Chr$(254)
    Err.Clear
End Property
Public Property Get Quote() As String
    On Error Resume Next
    ' value double quote
    Quote = Chr$(34)
    Err.Clear
End Property
Public Property Get Speed() As String
    On Error Resume Next
    ' get the speed of the modem
    Speed = mSpeed
    Err.Clear
End Property
Public Property Let Speed(nSpeed As String)
    On Error Resume Next
    ' set the speed of the modem, this will change settings
    mSpeed = nSpeed
    MSComm.Settings = nSpeed & ",n,8,1"
    PropertyChanged "Speed"
    PropertyChanged "Settings"
    Err.Clear
End Property
Public Property Get PortOpen() As Boolean
    On Error Resume Next
    ' get the status of the port
    PortOpen = mPortOpen
    Err.Clear
End Property
Public Property Let PortOpen(nPortOpen As Boolean)
    On Error GoTo ErrMsg
    ' set the status of the port
    MSComm.PortOpen = nPortOpen
    mPortOpen = nPortOpen
    PropertyChanged "PortOpen"
    Err.Clear
    Exit Property
ErrMsg:
    RaiseEvent Response(Err.Description)
    Err.Clear
End Property
Private Sub MSComm_OnComm()
    On Error Resume Next
    ' read contents of the data sent by modem
    Select Case MSComm.CommEvent
    Case comEvReceive
        m_Buffer = MSComm.Input
    End Select
    Err.Clear
End Sub
Public Function Connect(MyPort As String, MySpeed As String) As String
    On Error Resume Next
    ' connect to the gsm modem by specifying the port number and speed of the modem
    If MSComm.PortOpen = True Then MSComm.PortOpen = False
    Me.CommPort = Val(MyPort)
    Me.Speed = MySpeed
    MSComm.DTREnable = True
    MSComm.RTSEnable = True
    MSComm.RThreshold = 1
    MSComm.InBufferSize = 1024
    Me.PortOpen = True
    Me.Echo EchoOn
    Connect = IIf(Me.PortOpen = True, "OK", "ERROR")
    Err.Clear
End Function
Public Property Get ManufacturerIdentification() As String
    On Error Resume Next
    ' get the gadget manufacturer identification
    ManufacturerIdentification = Request("AT+GMI", , resultPos, vbCrLf)
    RaiseEvent Response(ManufacturerIdentification)
    Err.Clear
End Property
Public Property Get ModemSerialNumber() As String
    On Error Resume Next
    ' get the gadget modem serial/imei number
    ModemSerialNumber = Request("AT+GSN", , resultPos, vbCrLf)
    RaiseEvent Response(ModemSerialNumber)
    Err.Clear
End Property
Public Property Get RevisionIdentification() As String
    On Error Resume Next
    ' get the gadget revision identification
    Dim mAnswer As String
    mAnswer = Request("AT+GMR", , resultPos, vbCrLf)
    RevisionIdentification = MvField(mAnswer, 2, ":")
    RaiseEvent Response(RevisionIdentification)
    Err.Clear
End Property
Public Property Get ModelIdentification() As String
    On Error Resume Next
    ' get the gadget model indentification number
    ModelIdentification = Request("AT+GMM", , resultPos, vbCrLf)
    RaiseEvent Response(ModelIdentification)
    Err.Clear
End Property
Public Function SMS_NewMessageIndicate(bSet As Boolean) As String
    On Error Resume Next
    ' tell the gadget to notify computer of new sms received
    If bSet = True Then
        SMS_NewMessageIndicate = Request("AT+CNMI=1,1,0,0,0", , resultPos, vbCrLf)
    Else
        SMS_NewMessageIndicate = Request("AT+CNMI=0,0,0,0,0", , resultPos, vbCrLf)
    End If
    RaiseEvent Response(SMS_NewMessageIndicate)
    Err.Clear
End Function
Public Property Get SignalQualityMeasure() As String
    On Error Resume Next
    ' read the signal of the phone and return it as a percentage
    ' the maximum signal is 31
    Dim mAnswer As String
    mAnswer = Request("AT+CSQ", , resultPos, vbCrLf)
    mAnswer = MvField(mAnswer, 2, ":")
    mAnswer = MvField(mAnswer, 1, ",")
    mAnswer = (Val(mAnswer) / 31) * 100
    SignalQualityMeasure = Round(mAnswer, 0)
    RaiseEvent Response(SignalQualityMeasure)
    Err.Clear
End Property
Public Function PhoneBook_MemoryStorage(PhoneBookSelect As PhoneBookMemoryStorageEnum) As String
    On Error Resume Next
    ' select gadget phonebook memory
    Dim mAnswer As String
    Select Case PhoneBookSelect
    Case SimPhoneBook
        mAnswer = Request("AT+CPBS=" & Me.Quote & "SM" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = Request("AT+CPBS?", , resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
        mpbCapacity = MvField(mAnswer, 3, ",")
        mpbUsed = MvField(mAnswer, 2, ",")
        mpbMemory = MvField(mAnswer, 1, ",")
    Case MobileEquipmentPhoneBook
        mAnswer = Request("AT+CPBS=" & Me.Quote & "ME" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = Request("AT+CPBS?", , resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
        mpbCapacity = MvField(mAnswer, 3, ",")
        mpbUsed = MvField(mAnswer, 2, ",")
        mpbMemory = MvField(mAnswer, 1, ",")
    Case BothPhoneBooks
        mAnswer = Request("AT+CPBS=" & Me.Quote & "MT" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = Request("AT+CPBS?", , resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
        mpbCapacity = MvField(mAnswer, 3, ",")
        mpbUsed = MvField(mAnswer, 2, ",")
        mpbMemory = MvField(mAnswer, 1, ",")
    Case ReadPhoneBookSetting
        mAnswer = Request("AT+CPBS?", , resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
        mpbCapacity = MvField(mAnswer, 3, ",")
        mpbUsed = MvField(mAnswer, 2, ",")
        mpbMemory = MvField(mAnswer, 1, ",")
    Case SupportedMemory = 4
        'mAnswer = Request("AT+CPBS=?", , resultPos, vbCrLf)
        mAnswer = Request("AT+CPBS=?", , , vbCrLf)
    End Select
    PhoneBook_MemoryStorage = mAnswer
    RaiseEvent Response(PhoneBook_MemoryStorage)
    Err.Clear
End Function
Public Function PhoneBook_ReadEntry(Location As Long) As String
    On Error Resume Next
    ' read a specified phonebook entry using a location
    ' an empty location returns blank
    Dim mAnswer As String
    Dim mLoc As String
    Dim mNum As String
    Dim mNam As String
    mAnswer = Request("AT+CPBR=" & Location, , resultPos, vbCrLf)
    mAnswer = Replace$(mAnswer, Me.Quote, "")
    mAnswer = Trim$(Replace$(mAnswer, "+CPBR:", ""))
    Select Case LCase$(mAnswer)
    Case "not found", "+cme error: not found", ""
        mAnswer = ""
    Case Else
        mLoc = MvField(mAnswer, 1, ",")
        mNum = MvField(mAnswer, 2, ",")
        mNam = MvField(mAnswer, 4, ",")
        mAnswer = mLoc & "," & mNum & "," & mNam
    End Select
    PhoneBook_ReadEntry = mAnswer
    RaiseEvent Response(PhoneBook_ReadEntry)
    Err.Clear
End Function
Public Function PhoneBook_FindEntry(FullName As String) As String
    On Error Resume Next
    ' search for a phonebook entry and return result
    Dim mAnswer As String
    mAnswer = Request("AT+CPBF=" & Me.Quote & FullName & Me.Quote, , resultPos, vbCrLf)
    mAnswer = Replace$(mAnswer, Me.Quote, "")
    mAnswer = Trim$(Replace$(mAnswer, "+CPBF:", ""))
    Select Case LCase$(mAnswer)
    Case "not found", "", "+cme error: not found"
        mAnswer = "0"
    End Select
    PhoneBook_FindEntry = mAnswer
    RaiseEvent Response(PhoneBook_FindEntry)
    Err.Clear
End Function
Public Function PhoneBook_EntryExists(ByVal FullName As String, ByVal CellNo As String) As String
    On Error Resume Next
    ' search for a phonebook entry for the name and cellphone and return the location
    ' a zero location means not found
    Dim mAnswer As String
    Dim mPhone As String
    Dim mIndex As String
    mAnswer = Request("AT+CPBF=" & Me.Quote & FullName & Me.Quote, , resultPos, vbCrLf)
    mAnswer = Replace$(mAnswer, Me.Quote, "")
    mAnswer = Trim$(Replace$(mAnswer, "+CPBF:", ""))
    Select Case LCase$(mAnswer)
    Case "not found", "", "+cme error: not found"
        PhoneBook_EntryExists = "0"
    Case Else
        mIndex = MvField(mAnswer, 1, ",")
        mPhone = MvField(mAnswer, 2, ",")
        If mPhone = CellNo Then
            PhoneBook_EntryExists = mIndex
        Else
            PhoneBook_EntryExists = "0"
        End If
    End Select
    RaiseEvent Response(PhoneBook_EntryExists)
    Err.Clear
End Function
Public Function PhoneBook_WriteEntry(Location As Long, ByVal CellNo As String, ByVal FullName As String) As String
    On Error Resume Next
    ' write a phonebook entry at a particular location and refresh the count
    Dim mAnswer As String
    mAnswer = "AT+CPBW=" & Location & "," & Me.Quote & CellNo & Me.Quote & ",129," & Me.Quote & FullName & Me.Quote
    mAnswer = Request(mAnswer, , resultPos, vbCrLf)
    PhoneBook_WriteEntry = mAnswer
    If mAnswer = "OK" Then Me.PhoneBook_MemoryStorage (ReadPhoneBookSetting)
    RaiseEvent Response(PhoneBook_WriteEntry)
    Err.Clear
End Function
Public Function PhoneBook_DeleteEntry(Location As Long) As String
    On Error Resume Next
    ' delete a phonebook entry using the location
    Dim mAnswer As String
    mAnswer = Request("AT+CPBW=" & Location, , resultPos, vbCrLf)
    PhoneBook_DeleteEntry = mAnswer
    RaiseEvent Response(PhoneBook_DeleteEntry)
    Err.Clear
End Function
Public Property Get PhoneBook_AvailableIndexes() As String
    On Error Resume Next
    ' get the available phonebook indexes
    Dim mAnswer As String
    mAnswer = Request("AT+CPBR=?", , resultPos, vbCrLf)
    mAnswer = MvField(mAnswer, 2, ":")
    PhoneBook_AvailableIndexes = mAnswer
    RaiseEvent Response(PhoneBook_AvailableIndexes)
    Err.Clear
End Property
Public Function Echo(EchoStatus As EchoEnum) As Boolean
    On Error Resume Next
    ' turn echo off/on, off results in less traffic
    ' echo off returns the command with the result
    Dim mAnswer As String
    Select Case EchoStatus
    Case EchoOff
        mAnswer = Request("ATE0", , resultPos, vbCrLf)
    Case EchoOn
        mAnswer = Request("ATE1", , resultPos, vbCrLf)
    End Select
    If mAnswer = "OK" Then
        Echo = True
    Else
        Echo = False
    End If
    RaiseEvent Response(Echo)
    Err.Clear
End Function
Public Function Request(ByVal strCommand As String, Optional ByVal ExpectedResult As String = "", Optional Position As Long = -1, Optional ControlChars As String = vbCrLf) As String
    On Error GoTo ErrMsg
    ' send a request to the comm port and wait for the result
    ' the result will be delimited by chr(253)
    Dim sResult As String
    If Len(LogFile) > 0 Then
        FileUpdate LogFile, Now() & ", request received " & strCommand, "a"
    End If
    MSComm.Output = strCommand & ControlChars
    sResult = WaitReply(10, ExpectedResult)
    sResult = Replace$(sResult, vbNewLine, VM)
    If Len(LogFile) > 0 Then
        FileUpdate LogFile, Now() & ", response received " & sResult, "a"
    End If
    If Position > 0 Then
        sResult = MvField(sResult, Position, VM)
    Else
        sResult = sResult
    End If
    Request = DescriptiveError(sResult)
    Err.Clear
    Exit Function
ErrMsg:
    RaiseEvent Response(Err.Description)
    Err.Clear
End Function
Private Function WaitReply(lDelay As Long, WaitString As String) As String
    On Error Resume Next
    ' wait process for a data request from the
    ' ms comm port to be finalized
    Dim x As Long
    Dim bOK As Boolean
    DoEvents
    If WaitString = "" Then
        bOK = True
        For x = 1 To lDelay
            Sleep 500
            DoEvents
        Next
    Else
        bOK = False
        For x = 1 To lDelay
            DoEvents
            If InStr(m_Buffer, WaitString) Then
                bOK = True
                Exit For
            Else
                DoEvents
                Sleep 500
                DoEvents
            End If
        Next
    End If
    If Len(m_Buffer) > 0 Then
        DoEvents
    End If
    WaitReply = m_Buffer
    Err.Clear
End Function
Private Function MvField(ByVal strValue As String, Optional ByVal PartPosition As Long = 1, Optional ByVal Delimiter As String = ",", Optional TrimValue As Boolean = True) As String
    On Error Resume Next
    ' return a substring of a string delimited by a string specified
    Dim xResult As String
    Dim xArray() As String
    Dim xSize As Long
    If Len(strValue) = 0 Then Exit Function
    If InStr(1, strValue, Delimiter) = 0 Then
        MvField = strValue
        Err.Clear
        Exit Function
    End If
    xArray = Split(strValue, Delimiter)
    Select Case PartPosition
    Case -1
        PartPosition = UBound(xArray) + 1
    Case 0
        PartPosition = 1
    End Select
    xSize = UBound(xArray)
    If xSize = 0 Then
        MvField = ""
    Else
        xResult = xArray(PartPosition - 1)
        If TrimValue = True Then
            xResult = Trim$(xResult)
        End If
        MvField = xResult
    End If
    Err.Clear
End Function
Public Sub FileUpdate(ByVal filName As String, ByVal filLines As String, Optional ByVal Wora As String = "write")
    On Error Resume Next
    ' update contents of a file by either appending / creating a new entry
    Dim iFileNum As Long
    Dim cDir As String
    cDir = FileToken(filName, "p")
    CreateNestedDirectory cDir
    iFileNum = FreeFile
    Select Case LCase$(Left$(Wora, 1))
    Case "w"
        Open filName For Output As #iFileNum
        Case "a"
            Open filName For Append As #iFileNum
            End Select
            Print #iFileNum, filLines
        Close #iFileNum
        Err.Clear
End Sub
Private Function DescriptiveError(ByVal sResult As String) As String
    On Error Resume Next
    ' return descriptive phone error code
    Select Case sResult
    Case "+CME ERROR: 0":        DescriptiveError = "Phone failure"
    Case "+CME ERROR: 1":        DescriptiveError = "No connection to phone"
    Case "+CME ERROR: 2":        DescriptiveError = "Phone adapter link reserved"
    Case "+CME ERROR: 3":        DescriptiveError = "Operation not allowed"
    Case "+CME ERROR: 4":        DescriptiveError = "Operation not supported"
    Case "+CME ERROR: 5":        DescriptiveError = "PH_SIM PIN required"
    Case "+CME ERROR: 6":        DescriptiveError = "PH_FSIM PIN required"
    Case "+CME ERROR: 7":        DescriptiveError = "PH_FSIM PUK required"
    Case "+CME ERROR: 10":        DescriptiveError = "SIM not inserted"
    Case "+CME ERROR: 11":        DescriptiveError = "SIM PIN required"
    Case "+CME ERROR: 12":        DescriptiveError = "SIM PUK required"
    Case "+CME ERROR: 13":        DescriptiveError = "SIM failure"
    Case "+CME ERROR: 14":        DescriptiveError = "SIM busy"
    Case "+CME ERROR: 15":        DescriptiveError = "SIM wrong"
    Case "+CME ERROR: 16":        DescriptiveError = "Incorrect password"
    Case "+CME ERROR: 17":        DescriptiveError = "SIM PIN2 required"
    Case "+CME ERROR: 18":        DescriptiveError = "SIM PUK2 required"
    Case "+CME ERROR: 20":        DescriptiveError = "Memory full"
    Case "+CME ERROR: 21":        DescriptiveError = "Invalid index"
    Case "+CME ERROR: 22":        DescriptiveError = "Not found"
    Case "+CME ERROR: 23":        DescriptiveError = "Memory failure"
    Case "+CME ERROR: 24":        DescriptiveError = "Text string too long"
    Case "+CME ERROR: 25":        DescriptiveError = "Invalid characters in text string"
    Case "+CME ERROR: 26":        DescriptiveError = "Dial string too long"
    Case "+CME ERROR: 27":        DescriptiveError = "Invalid characters in dial string"
    Case "+CME ERROR: 30":        DescriptiveError = "No network service"
    Case "+CME ERROR: 31":        DescriptiveError = "Network timeout"
    Case "+CME ERROR: 32":        DescriptiveError = "Network not allowed, emergency calls only"
    Case "+CME ERROR: 40":        DescriptiveError = "Network personalization PIN required"
    Case "+CME ERROR: 41":        DescriptiveError = "Network personalization PUK required"
    Case "+CME ERROR: 42":        DescriptiveError = "Network subset personalization PIN required"
    Case "+CME ERROR: 43":        DescriptiveError = "Network subset personalization PUK required"
    Case "+CME ERROR: 44":        DescriptiveError = "Service provider personalization PIN required"
    Case "+CME ERROR: 45":        DescriptiveError = "Service provider personalization PUK required"
    Case "+CME ERROR: 46":        DescriptiveError = "Corporate personalization PIN required"
    Case "+CME ERROR: 47":        DescriptiveError = "Corporate personalization PUK required"
    Case "+CME ERROR: 48":        DescriptiveError = "PH-SIM PUK required"
    Case "+CME ERROR: 100":        DescriptiveError = "Unknown error"
    Case "+CME ERROR: 103":        DescriptiveError = "Illegal MS"
    Case "+CME ERROR: 106":        DescriptiveError = "Illegal ME"
    Case "+CME ERROR: 107":        DescriptiveError = "GPRS services not allowed"
    Case "+CME ERROR: 111":        DescriptiveError = "PLMN not allowed"
    Case "+CME ERROR: 112":        DescriptiveError = "Location area not allowed"
    Case "+CME ERROR: 113":        DescriptiveError = "Roaming not allowed in this location area"
    Case "+CME ERROR: 126":        DescriptiveError = "Operation temporary not allowed"
    Case "+CME ERROR: 132":        DescriptiveError = "Service operation not supported"
    Case "+CME ERROR: 133":        DescriptiveError = "Requested service option not subscribed"
    Case "+CME ERROR: 134":        DescriptiveError = "Service option temporary out of order"
    Case "+CME ERROR: 148":        DescriptiveError = "Unspecified GPRS error"
    Case "+CME ERROR: 149":        DescriptiveError = "PDP authentication failure"
    Case "+CME ERROR: 150":        DescriptiveError = "Invalid mobile class"
    Case "+CME ERROR: 256":        DescriptiveError = "Operation temporarily not allowed"
    Case "+CME ERROR: 257":        DescriptiveError = "Call barred"
    Case "+CME ERROR: 258":        DescriptiveError = "Phone is busy"
    Case "+CME ERROR: 259":        DescriptiveError = "User abort"
    Case "+CME ERROR: 260":        DescriptiveError = "Invalid dial string"
    Case "+CME ERROR: 261":        DescriptiveError = "SS not executed"
    Case "+CME ERROR: 262":        DescriptiveError = "SIM Blocked"
    Case "+CME ERROR: 263":        DescriptiveError = "Invalid block"
    Case "+CME ERROR: 772":        DescriptiveError = "SIM powered down"
    Case "+CMS ERROR: 1": DescriptiveError = "Unassigned (unallocated) number"
    Case "+CMS ERROR: 8": DescriptiveError = "Operator determined barring"
    Case "+CMS ERROR: 10": DescriptiveError = "Call barred"
    Case "+CMS ERROR: 21": DescriptiveError = "Short message transfer rejected"
    Case "+CMS ERROR: 27": DescriptiveError = "Destination out of service"
    Case "+CMS ERROR: 28": DescriptiveError = "Unidentified subscriber"
    Case "+CMS ERROR: 29": DescriptiveError = "Facility rejected"
    Case "+CMS ERROR: 30": DescriptiveError = "Unknown subscriber"
    Case "+CMS ERROR: 38": DescriptiveError = "Network out of order"
    Case "+CMS ERROR: 41": DescriptiveError = "Temporary failure"
    Case "+CMS ERROR: 42": DescriptiveError = "Congestion"
    Case "+CMS ERROR: 47": DescriptiveError = "Resources unavailable, unspecified"
    Case "+CMS ERROR: 50": DescriptiveError = "Requested facility not subscribed"
    Case "+CMS ERROR: 69": DescriptiveError = "Requested facility not implemented"
    Case "+CMS ERROR: 81": DescriptiveError = "Invalid short message transfer reference value"
    Case "+CMS ERROR: 95": DescriptiveError = "Invalid message, unspecified"
    Case "+CMS ERROR: 96": DescriptiveError = "Invalid mandatory information"
    Case "+CMS ERROR: 97": DescriptiveError = "Message type non-existent or not implemented"
    Case "+CMS ERROR: 98": DescriptiveError = "Message not compatible with short message protocol state"
    Case "+CMS ERROR: 99": DescriptiveError = "Information element non-existent or not implemented"
    Case "+CMS ERROR: 111": DescriptiveError = "Protocol error, unspecified"
    Case "+CMS ERROR: 127": DescriptiveError = "Interworking, unspecified"
    Case "+CMS ERROR: 128": DescriptiveError = "Telematic interworking not supported"
    Case "+CMS ERROR: 129": DescriptiveError = "Short message Type 0 not supported"
    Case "+CMS ERROR: 130": DescriptiveError = "Cannot replace short message"
    Case "+CMS ERROR: 143": DescriptiveError = "Unspecified TP-PID error"
    Case "+CMS ERROR: 144": DescriptiveError = "Data coding scheme (alphabet) not supported"
    Case "+CMS ERROR: 145": DescriptiveError = "Message class not supported"
    Case "+CMS ERROR: 159": DescriptiveError = "Unspecified TP-DCS error"
    Case "+CMS ERROR: 160": DescriptiveError = "Command cannot be actioned"
    Case "+CMS ERROR: 161": DescriptiveError = "Command unsupported"
    Case "+CMS ERROR: 175": DescriptiveError = "Unspecified TP-Command error"
    Case "+CMS ERROR: 176": DescriptiveError = "TPDU not supported"
    Case "+CMS ERROR: 192": DescriptiveError = "SC busy"
    Case "+CMS ERROR: 193": DescriptiveError = "No SC subscription"
    Case "+CMS ERROR: 194": DescriptiveError = "SC system failure"
    Case "+CMS ERROR: 195": DescriptiveError = "Invalid SME address"
    Case "+CMS ERROR: 196": DescriptiveError = "Destination SME barred"
    Case "+CMS ERROR: 197": DescriptiveError = "SM Rejected-Duplicate SM"
    Case "+CMS ERROR: 198": DescriptiveError = "TP-VPF not supported"
    Case "+CMS ERROR: 199": DescriptiveError = "TP-VP not supported"
    Case "+CMS ERROR: 208": DescriptiveError = "SIM SMS storage full"
    Case "+CMS ERROR: 209": DescriptiveError = "No SMS storage capability in SIM"
    Case "+CMS ERROR: 210": DescriptiveError = "Error in MS"
    Case "+CMS ERROR: 211": DescriptiveError = "Memory Capacity Exceeded"
    Case "+CMS ERROR: 212": DescriptiveError = "SIM Application Toolkit Busy"
    Case "+CMS ERROR: 255": DescriptiveError = "Unspecified error cause"
    Case "+CMS ERROR: 300": DescriptiveError = "ME failure"
    Case "+CMS ERROR: 301": DescriptiveError = "SMS service of ME reserved"
    Case "+CMS ERROR: 302": DescriptiveError = "Operation not allowed"
    Case "+CMS ERROR: 303": DescriptiveError = "Operation not supported"
    Case "+CMS ERROR: 304": DescriptiveError = "Invalid PDU mode parameter"
    Case "+CMS ERROR: 305": DescriptiveError = "Invalid text mode parameter"
    Case "+CMS ERROR: 310": DescriptiveError = "SIM not inserted"
    Case "+CMS ERROR: 311": DescriptiveError = "SIM PIN required"
    Case "+CMS ERROR: 312": DescriptiveError = "PH-SIM PIN required"
    Case "+CMS ERROR: 313": DescriptiveError = "SIM failure"
    Case "+CMS ERROR: 314": DescriptiveError = "SIM busy"
    Case "+CMS ERROR: 315": DescriptiveError = "SIM wrong"
    Case "+CMS ERROR: 316": DescriptiveError = "SIM PUK required"
    Case "+CMS ERROR: 317": DescriptiveError = "SIM PIN2 required"
    Case "+CMS ERROR: 318": DescriptiveError = "SIM PUK2 required"
    Case "+CMS ERROR: 320": DescriptiveError = "memory failure"
    Case "+CMS ERROR: 321": DescriptiveError = "Invalid memory index"
    Case "+CMS ERROR: 322": DescriptiveError = "Memory full"
    Case "+CMS ERROR: 330": DescriptiveError = "SMSC address unknown"
    Case "+CMS ERROR: 331": DescriptiveError = "No network service"
    Case "+CMS ERROR: 332": DescriptiveError = "Network timeout"
    Case "+CMS ERROR: 340": DescriptiveError = "No +CNMA acknowledgement expected"
    Case "+CMS ERROR: 17": DescriptiveError = "Network failure"
    Case "+CMS ERROR: 22": DescriptiveError = "Congestion"
    Case "+CMS ERROR: 500": DescriptiveError = "Unknown Error"
    Case Else
        DescriptiveError = sResult
    End Select
    Err.Clear
End Function
Public Property Get PhoneBook_AvailableIndex() As Long
    On Error Resume Next
    ' get next available index
    Dim rsCnt As Long
    Dim phEntry As String
    Dim pCapacity As Long
    Call Me.PhoneBook_MemoryStorage(ReadPhoneBookSetting)
    pCapacity = Me.PhoneBook_Capacity
    PhoneBook_AvailableIndex = -1
    For rsCnt = 1 To pCapacity
        ' read entry at specified index
        phEntry = Me.PhoneBook_ReadEntry(rsCnt)
        ' if successfull return index,cellno,fullname
        If Len(phEntry) = 0 Then
            PhoneBook_AvailableIndex = rsCnt
            Exit For
        End If
        DoEvents
    Next
    RaiseEvent Response(PhoneBook_AvailableIndex)
    Err.Clear
End Property
Public Function PhoneBook_AddEntry(ByVal sNumber As String, ByVal sName As String) As String
    On Error Resume Next
    'add an entry to the phonebook
    Dim availableIndex As Long
    availableIndex = Me.PhoneBook_AvailableIndex
    If availableIndex = -1 Then
        PhoneBook_AddEntry = "Phonebook Full"
    Else
        PhoneBook_AddEntry = Me.PhoneBook_WriteEntry(availableIndex, sNumber, sName)
    End If
    RaiseEvent Response(PhoneBook_AddEntry)
    Err.Clear
End Function
Public Function SMS_MessageFormat(MessageFormatAction As MessageFormatEnum) As String
    On Error Resume Next
    ' message format management
    Dim mAnswer As String
    Select Case MessageFormatAction
    Case TextFormat
        SMS_MessageFormat = Request("AT+CMGF=1", , resultPos, vbCrLf)
    Case PDUFormat
        SMS_MessageFormat = Request("AT+CMGF=0", , resultPos, vbCrLf)
    Case ReadFormat
        mAnswer = Request("AT+CMGF?", , resultPos, vbCrLf)
        mAnswer = Trim$(MvField(mAnswer, 2, ":"))
        Select Case mAnswer
        Case 0
            SMS_MessageFormat = "PDU"
        Case 1
            SMS_MessageFormat = "TEXT"
        End Select
    End Select
    RaiseEvent Response(SMS_MessageFormat)
    Err.Clear
End Function
Public Function SMS_MemoryStorage(SelectMemory As SmsMemoryStorageEnum) As String
    On Error Resume Next
    ' select phonebook memory
    Dim mAnswer As String
    Select Case SelectMemory
    Case SimMemory
        mAnswer = Request("AT+CPMS=" & Me.Quote & "SM" & Me.Quote & "," & Me.Quote & "SM" & Me.Quote & "," & Me.Quote & "SM" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "SM"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "SM"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "SM"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case MobileEquipmentMemory
        mAnswer = Request("AT+CPMS=" & Me.Quote & "ME" & Me.Quote & "," & Me.Quote & "ME" & Me.Quote & "," & Me.Quote & "ME" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "ME"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "ME"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "ME"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case BothMemories
        mAnswer = Request("AT+CPMS=" & Me.Quote & "MT" & Me.Quote & "," & Me.Quote & "MT" & Me.Quote & "," & Me.Quote & "MT" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "MT"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "MT"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "MT"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case ReadMemorySetting
        mAnswer = Request("AT+CPMS?", "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
        mReadDelete = MvField(mAnswer, 1, ",")
        mReadDeleteUsed = Val(MvField(mAnswer, 2, ","))
        mReadDeleteCapacity = Val(MvField(mAnswer, 3, ","))
        mWriteSend = MvField(mAnswer, 4, ",")
        mWriteSendUsed = Val(MvField(mAnswer, 5, ","))
        mWriteSendCapacity = Val(MvField(mAnswer, 6, ","))
        mReceived = MvField(mAnswer, 7, ",")
        mReceivedUsed = Val(MvField(mAnswer, 8, ","))
        mReceivedCapacity = Val(MvField(mAnswer, 9, ","))
        mAnswer = "OK"
    Case BroadcastMessageStorage
        mAnswer = Request("AT+CPMS=" & Me.Quote & "BM" & Me.Quote & "," & Me.Quote & "BM" & Me.Quote & "," & Me.Quote & "BM" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "BM"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "BM"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "BM"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case StatusReportStorage
        mAnswer = Request("AT+CPMS=" & Me.Quote & "SR" & Me.Quote & "," & Me.Quote & "SR" & Me.Quote & "," & Me.Quote & "SR" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "SR"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "SR"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "SR"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    Case TerminalAdapterStorage
        mAnswer = Request("AT+CPMS=" & Me.Quote & "TA" & Me.Quote & "," & Me.Quote & "TA" & Me.Quote & "," & Me.Quote & "TA" & Me.Quote, "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        Select Case mAnswer
        Case "ERROR"
            mReadDelete = ""
            mReadDeleteUsed = -1
            mReadDeleteCapacity = -1
            mWriteSend = ""
            mWriteSendUsed = -1
            mWriteSendCapacity = -1
            mReceived = ""
            mReceivedUsed = -1
            mReceivedCapacity = -1
        Case Else
            mReadDelete = "TA"
            mReadDeleteUsed = Val(MvField(mAnswer, 1, ","))
            mReadDeleteCapacity = Val(MvField(mAnswer, 2, ","))
            mWriteSend = "TA"
            mWriteSendUsed = Val(MvField(mAnswer, 3, ","))
            mWriteSendCapacity = Val(MvField(mAnswer, 4, ","))
            mReceived = "TA"
            mReceivedUsed = Val(MvField(mAnswer, 5, ","))
            mReceivedCapacity = Val(MvField(mAnswer, 6, ","))
            mAnswer = "OK"
        End Select
    End Select
    SMS_MemoryStorage = mAnswer
    RaiseEvent Response(SMS_MemoryStorage)
    Err.Clear
End Function
Public Function SMS_CentreNumber(CentreNumberAction As CentreNumberEnum, Optional SMSC As String = "") As String
    On Error Resume Next
    ' centre number management
    Dim mAnswer As String
    Select Case CentreNumberAction
    Case SetCentreNumber
        mAnswer = Request("AT+CSCA=" & Me.Quote & SMSC & Me.Quote, "", resultPos, vbCrLf)
    Case ReadCentreNumber
        mAnswer = Request("AT+CSCA?", "", resultPos, vbCrLf)
        mAnswer = MvField(mAnswer, 2, ":")
        mAnswer = MvField(mAnswer, 1, ",")
        mAnswer = Replace$(mAnswer, Me.Quote, "")
    End Select
    SMS_CentreNumber = mAnswer
    RaiseEvent Response(SMS_CentreNumber)
    Err.Clear
End Function
Public Property Get SMS_ReadDeleteStorage() As String
    On Error Resume Next
    SMS_ReadDeleteStorage = mReadDelete
    Err.Clear
End Property
Public Property Get SMS_ReadDeleteStorageUsed() As Long
    On Error Resume Next
    SMS_ReadDeleteStorageUsed = mReadDeleteUsed
    Err.Clear
End Property
Public Property Get SMS_ReadDeleteStorageCapacity() As Long
    On Error Resume Next
    SMS_ReadDeleteStorageCapacity = mReadDeleteCapacity
    Err.Clear
End Property
Public Property Get SMS_WriteSendStorage() As String
    On Error Resume Next
    SMS_WriteSendStorage = mWriteSend
    Err.Clear
End Property
Public Property Get SMS_WriteSendStorageUsed() As Long
    On Error Resume Next
    SMS_WriteSendStorageUsed = mWriteSendUsed
    Err.Clear
End Property
Public Property Get SMS_WriteSendStorageCapacity() As Long
    On Error Resume Next
    SMS_WriteSendStorageCapacity = mWriteSendCapacity
    Err.Clear
End Property
Public Property Get SMS_ReceivedStorage() As String
    On Error Resume Next
    SMS_ReceivedStorage = mReceived
    Err.Clear
End Property
Public Property Get SMS_ReceivedStorageUsed() As Long
    On Error Resume Next
    SMS_ReceivedStorageUsed = mReceivedUsed
    Err.Clear
End Property
Public Property Get SMS_ReceivedStorageCapacity() As Long
    On Error Resume Next
    SMS_ReceivedStorageCapacity = mReceivedCapacity
    Err.Clear
End Property
Public Property Get PhoneBook_Used() As Long
    On Error Resume Next
    PhoneBook_Used = mpbUsed
    Err.Clear
End Property
Public Property Get PhoneBook_Capacity() As Long
    On Error Resume Next
    PhoneBook_Capacity = mpbCapacity
    Err.Clear
End Property
Public Property Get PhoneBook_Memory() As String
    On Error Resume Next
    ' update property
    PhoneBook_Memory = mpbMemory
    Err.Clear
End Property
Public Property Get SubscriberNumber() As String
    On Error Resume Next
    ' get the subscriber number
    SubscriberNumber = Request("AT+CNUM", , resultPos, vbCrLf)
    RaiseEvent Response(SubscriberNumber)
    Err.Clear
End Property
Public Property Get InternationalMobileSubscriberIdentity() As String
    On Error Resume Next
    ' get the international mobile subscriber identity
    InternationalMobileSubscriberIdentity = Request("AT+CIMI", , resultPos, vbCrLf)
    RaiseEvent Response(InternationalMobileSubscriberIdentity)
    Err.Clear
End Property
Public Function SMS_ReadMessageEntry(msgIndex As Long) As String
    On Error Resume Next
    ' read the message stored at location
    Dim mAnswer As String
    Dim tmpAns As String
    Dim MsgType As String
    Dim msgFrom As String
    Dim msgDate As String
    Dim msgTime As String
    Dim msgContents As String
    Dim pPos As Long
    mAnswer = Request("AT+CMGR=" & msgIndex, , , vbCrLf)
    mAnswer = MvRest(mAnswer, 2, Me.VM)
    mAnswer = Trim$(Replace$(mAnswer, "+CMGR:", ""))
    tmpAns = MvField(mAnswer, 1, VM)
    Select Case tmpAns
    Case "OK"
        mAnswer = ""
    Case Else
        MsgType = MvField(mAnswer, 1, ",")
        MsgType = Replace$(MsgType, Quote, "")
        msgFrom = MvField(mAnswer, 2, ",")
        msgFrom = Replace$(msgFrom, Quote, "")
        msgDate = MvField(mAnswer, 4, ",")
        msgDate = Replace$(msgDate, Quote, "")
        msgDate = FixSmsDate(msgDate)
        msgTime = MvField(mAnswer, 5, ",")
        msgTime = MvField(msgTime, 1, VM)
        msgTime = Replace$(msgTime, Quote, "")
        pPos = InStr(1, msgTime, "+")
        If pPos > 0 Then
            msgTime = Left$(msgTime, pPos - 1)
        End If
        msgContents = MvRest(mAnswer, 2, VM)
        msgContents = Replace$(msgContents, VM & VM & "OK" & VM, "")
        msgContents = RemAllVM(msgContents)
        mAnswer = msgIndex & FM & MsgType & FM & msgFrom & FM & msgDate & " " & msgTime & FM & msgContents
    End Select
    SMS_ReadMessageEntry = mAnswer
    RaiseEvent Response(SMS_ReadMessageEntry)
    Err.Clear
End Function
Public Sub SMS_ListView(progBar As Variant, LstView As Variant, Optional mIcon As String = "", Optional mSmallIcon As String = "", Optional MsgType As SmsTypesEnum)
    On Error Resume Next
    ' load contents of the selected message store messages
    ' to the listview
    Dim rsCnt As Long
    Dim phEntry As String
    Dim lstItem As Variant
    Dim spLine() As String
    Dim nCollection As New Collection
    Set nCollection = SMS_ReadMessages
    progBar.Max = nCollection.Count
    progBar.Min = 0
    progBar.Value = 0
    ' create headings
    LstViewMakeHeadings LstView, "Msg ID,Cellphone No.,Contents,Time"
    ' loop through the messages starting from location 1 to the full capacity of the phone
    For rsCnt = 1 To nCollection.Count
        progBar.Value = rsCnt
        ' read entry at specified index
        phEntry = nCollection(rsCnt)
        ' if successfull return msg index, type,cellnumber,date,message
        If Len(phEntry) > 0 Then
            spLine = Split(phEntry, FM)
            Select Case MsgType
            Case Rec
                If spLine(1) = "REC READ" Or spLine(1) = "REC UNREAD" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case RecRead
                If spLine(1) = "REC READ" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case RecUnread
                If spLine(1) = "REC UNREAD" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case StoSent
                If spLine(1) = "STO SENT" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            Case StoUnsent
                If spLine(1) = "STO UNSENT" Then
                    GoTo AddLine
                Else
                    GoTo NextLine
                End If
            End Select
AddLine:
            Set lstItem = LstView.ListItems.Add(, , spLine(0))
            lstItem.SubItems(1) = spLine(2)
            lstItem.SubItems(2) = spLine(4)
            lstItem.SubItems(3) = spLine(3)
            If Len(mIcon) > 0 Then lstItem.Icon = mIcon
            If Len(mSmallIcon) > 0 Then lstItem.SmallIcon = mSmallIcon
        End If
NextLine:
        DoEvents
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Function MvRest(ByVal strData As String, Optional ByVal startPos As Long = 1, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    ' get the string from a substring position to the end of the
    ' delimited string
    Dim spData() As String
    Dim spCnt As Long
    Dim intLoop As Long
    Dim strL As String
    Dim strM As String
    MvRest = ""
    strM = ""
    If Len(Delim) = 0 Then Delim = Me.VM
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    spData = Split(strData, Delim)
    spCnt = UBound(spData)
    Select Case startPos
    Case -1
        MvRest = Trim$(spData(spCnt))
    Case Else
        strL = ""
        startPos = startPos - 1
        For intLoop = startPos To spCnt
            strL = spData(intLoop)
            If intLoop = spCnt Then
                strM = strM & strL
            Else
                strM = strM & strL & Delim
            End If
        Next
        MvRest = strM
    End Select
    Err.Clear
End Function
Public Function SMS_DeleteEntry(Location As Long) As String
    On Error Resume Next
    ' delete a sms and refresh if ok
    Dim mAnswer As String
    mAnswer = Request("AT+CMGD=" & Location, , resultPos, vbCrLf)
    SMS_DeleteEntry = mAnswer
    If mAnswer = "OK" Then Call SMS_MemoryStorage(ReadMemorySetting)
    RaiseEvent Response(SMS_DeleteEntry)
    Err.Clear
End Function
Private Function SMS_ReadMessages() As Collection
    On Error Resume Next
    ' read all sms messages from selected memory
    Dim mAnswer As String
    Dim nCollection As New Collection
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsLines() As String
    Dim msgIndex As String
    Dim MsgType As String
    Dim msgFrom As String
    Dim msgDate As String
    Dim msgTime As String
    Dim msgContents As String
    Dim pPos As Long
    mAnswer = Me.SMS_MessageFormat(TextFormat)
    If mAnswer = "OK" Then
        mAnswer = Me.Request("AT+CMGL=" & Quote & "ALL" & Quote, , , vbCrLf)
    End If
    rsLines = Split(mAnswer, "+CMGL:")
    rsTot = UBound(rsLines)
    For rsCnt = 0 To rsTot
        rsLines(rsCnt) = Trim$(rsLines(rsCnt))
        rsStr = rsLines(rsCnt)
        msgIndex = MvField(rsStr, 1, ",")
        If IsNumeric(msgIndex) = True Then
            MsgType = MvField(rsStr, 2, ",")
            MsgType = Replace$(MsgType, Quote, "")
            msgFrom = MvField(rsStr, 3, ",")
            msgFrom = Replace$(msgFrom, Quote, "")
            If msgFrom = "6" Then
                ' report
                GoTo NextRecord
            End If
            msgDate = MvField(rsStr, 5, ",")
            msgDate = Replace$(msgDate, Quote, "")
            msgDate = FixSmsDate(msgDate)
            msgTime = MvField(rsStr, 6, ",")
            msgTime = MvField(msgTime, 1, VM)
            msgTime = Replace$(msgTime, Quote, "")
            pPos = InStr(1, msgTime, "+")
            If pPos > 0 Then
                msgTime = Left$(msgTime, pPos - 1)
            End If
            msgContents = MvRest(rsStr, 2, VM)
            msgContents = Replace$(msgContents, VM & VM & "OK" & VM, "")
            msgContents = RemAllVM(msgContents)
            rsStr = msgIndex & FM & MsgType & FM & msgFrom & FM & msgDate & " " & msgTime & FM & msgContents
            nCollection.Add rsStr
        End If
NextRecord:
    Next
    Set SMS_ReadMessages = nCollection
    Err.Clear
End Function
Private Function SMS_SendSmall(sNumber As String, sMessage As String) As String
    On Error Resume Next
    ' send an sms to a phone, set textmode format just in case
    Dim mAnswer As String
    mAnswer = Request("AT+CMGS=" & Quote & sNumber & Quote, , resultPos, vbCr)
    Select Case mAnswer
    Case ">"
        mAnswer = Request(sMessage, , 4, Chr$(26))
    Case Else
        mAnswer = "ERROR"
    End Select
    SMS_SendSmall = mAnswer
    RaiseEvent Response(SMS_SendSmall)
    Err.Clear
End Function
Public Function SMS_Send(sNumber As String, sMessage As String) As String
    On Error Resume Next
    ' send an sms to a phone, set textmode format just in case
    Dim mAnswer As String
    Dim msgSent As Long
    Dim msgOne As Long
    Dim msgTwo As Long
    msgOne = 0
    msgTwo = 0
    mAnswer = SMS_MessageFormat(TextFormat)
    If mAnswer = "OK" Then
        If Len(sMessage) > 160 Then
            mAnswer = SMS_SendSmall(sNumber, Mid$(sMessage, 1, 160))
            Select Case LCase$(mAnswer)
            Case "ok"
                msgOne = 0
            Case Else
                msgOne = 1
            End Select
            mAnswer = SMS_SendSmall(sNumber, Mid$(sMessage, 161, 160))
            Select Case LCase$(mAnswer)
            Case "ok"
                msgTwo = 0
            Case Else
                msgTwo = 1
            End Select
            msgSent = msgOne + msgTwo
            If msgSent = 0 Then
                mAnswer = "OK"
            Else
                mAnswer = "ERROR"
            End If
        Else
            mAnswer = SMS_SendSmall(sNumber, sMessage)
        End If
    Else
        mAnswer = "ERROR"
    End If
    SMS_Send = mAnswer
    RaiseEvent Response(SMS_Send)
    Err.Clear
End Function
Private Function RemAllVM(ByVal StrString As String) As String
    On Error Resume Next
    ' remove trailing delimiters
    Dim strSize As Long
    Dim strLast As String
    Dim tmpstring As String
    tmpstring = StrString
    strLast = Right$(tmpstring, 1)
    Do While strLast = VM
        strSize = Len(tmpstring) - 1
        tmpstring = Left$(tmpstring, strSize)
        strLast = Right$(tmpstring, 1)
    Loop
    RemAllVM = tmpstring
    Err.Clear
End Function
Public Function PhoneBook_Export(progBar As Variant, exportFile As String) As Boolean
    On Error Resume Next
    ' load contents of the selected phonebook
    ' to the listview
    Dim rsCnt As Long
    Dim phEntry As String
    'Dim lstItem As ListItem
    'Dim spLine() As String
    'Dim lstTotal As Long
    Dim pUsed As Long
    Dim pCapacity As Long
    Dim pWritten As Long
    pWritten = 0
    If FileExists(exportFile) = True Then Kill exportFile
    phEntry = PhoneBook_MemoryStorage(ReadPhoneBookSetting)
    If phEntry = "ERROR" Then
        PhoneBook_Export = False
    Else
        pCapacity = PhoneBook_Capacity
        pUsed = PhoneBook_Used
        progBar.Max = pCapacity
        progBar.Min = 0
        progBar.Value = 0
        FileUpdate exportFile, "Index,Cellphone No,Full Name", "a"
        For rsCnt = 1 To pCapacity
            progBar.Value = rsCnt
            ' read entry at specified index
            phEntry = Me.PhoneBook_ReadEntry(rsCnt)
            ' if successfull return index,cellno,fullname
            If Len(phEntry) > 0 Then
                pWritten = pWritten + 1
                FileUpdate exportFile, phEntry, "a"
            End If
            ' ensure that if we have reached the used limit, exit the loop
            ' we do not want to read the empty contacts anymore
            If pWritten = pUsed Then Exit For
            DoEvents
        Next
    End If
    progBar.Value = 0
    PhoneBook_Export = True
    Err.Clear
End Function
Private Function FileExists(ByVal Filename As String) As Boolean
    On Error Resume Next
    ' returns the existance of a file
    FileExists = False
    If Len(Filename) = 0 Then
        Err.Clear
        Exit Function
    End If
    FileExists = IIf(Dir$(Filename) <> "", True, False)
    Err.Clear
End Function
Public Function PhoneBook_Import(progBar As Variant, importFile As String) As Long
    On Error Resume Next
    ' import phonebook details to selected memory
    Dim phEntry As String
    Dim pWritten As Long
    Dim fData As String
    Dim fLines() As String
    Dim fTot As Long
    Dim fCnt As Long
    Dim cName As String
    Dim cNumber As String
    Dim cIndex As String
    Dim cLocation As String
    Dim cResult As String
    fData = FileData(importFile)
    fLines = Split(fData, vbNewLine)
    fTot = UBound(fLines)
    progBar.Max = fTot + 1
    progBar.Min = 0
    progBar.Value = 0
    pWritten = 0
    For fCnt = 0 To fTot + 1
        progBar.Value = fCnt + 1
        phEntry = fLines(fCnt)
        If Len(phEntry) = 0 Then GoTo NextRow
        cIndex = MvField(phEntry, 1, ",")
        cNumber = MvField(phEntry, 2, ",")
        cName = MvField(phEntry, 3, ",")
        If cIndex = "Index" Then GoTo NextRow
        cLocation = Me.PhoneBook_EntryExists(cName, cNumber)
        If cLocation = "0" Then
            cResult = PhoneBook_AddEntry(cNumber, cName)
            If cResult = "OK" Then
                pWritten = pWritten + 1
            End If
        End If
NextRow:
        DoEvents
    Next
    PhoneBook_Import = pWritten
    Err.Clear
End Function
Private Function FileData(ByVal Filename As String) As String
    On Error Resume Next
    ' return contents of file
    Dim sLen As Long
    Dim fileNum As Long
    Dim Size As Long
    fileNum = FreeFile
    Size = FileLen(Filename)
    Open Filename For Input Access Read As #fileNum
        sLen = LOF(fileNum)
        FileData = Input(sLen, #fileNum)
    Close #fileNum
    Err.Clear
End Function
Private Function Biner(Bilangan) As String
    On Error Resume Next
    ' binary management, used in pdu mode
    Dim Basis As Integer
    Dim Hsltemp As Variant
    Dim sisa As Variant
    Dim HslBagi As Variant
    Hsltemp = ""
    sisa = ""
    Basis = 2
    Do
        Hsltemp = sisa & Hsltemp
        HslBagi = Bilangan \ Basis
        sisa = Bilangan Mod Basis
        Bilangan = HslBagi
    Loop Until HslBagi <= 1
    Biner = HslBagi & sisa & Hsltemp
    Biner = Right("0000000" & Biner, 7)
    Err.Clear
End Function
Private Function CharHex(ByVal Txt As String, ByVal bit As Integer)
    On Error Resume Next
    'Create 7 bit / 8 bit
    'charhex(str,7)-->7 bit for receiving SMS
    'charhex(str,8)-->8 bit for send
    ' used in pdu mode
    Dim i As Integer
    Dim bin As String
    Dim nbin As String
    Dim n As String
    Dim bil As Integer
    Dim sisa As Integer
    Dim lbin As Integer
    Dim nol As String
    bin = ""
    nbin = ""
    If bit = 7 Then
        For i = 1 To Len(Txt) Step 2
            n = Mid(Txt, i, 2)
            bin = HexToBin(n) & bin
        Next
        bil = Len(bin) \ bit
        sisa = Len(bin) Mod bit
        For i = 1 To (Len(bin) - sisa) Step bit
            nbin = Chr$(HexToDec(BinToHex(Mid(bin, i + sisa, bit)))) & nbin
        Next
    Else
        For i = 1 To Len(Txt)
            n = Mid(Txt, i, 1)
            bin = Biner(Asc(n)) & bin
        Next
        sisa = Len(bin) Mod bit
        If sisa > 0 Then
            For i = 1 To bit - sisa
                nol = nol & "0"
            Next
        End If
        bin = nol & bin
        bil = Len(bin) \ bit
        For i = 1 To bil
            nbin = nbin & BinToHex(Mid(bin, Len(bin) + 1 - bit * i, bit))
        Next
    End If
    CharHex = nbin
    Err.Clear
End Function
Private Function BinToHex(ByVal Biner As String) As String
    On Error Resume Next
    ' bin to hex, used in pdu mode
    Dim bin As String
    Dim n As String
    Dim nil As String
    Dim i As Integer
    bin = ""
    Biner = Right("00000000" & Biner, 8)
    For i = 1 To 2
        bin = Mid(Biner, Len(Biner) + 1 - 4 * i, 4)
        Select Case bin
        Case "0000": n = "0"
        Case "0001": n = "1"
        Case "0010": n = "2"
        Case "0011": n = "3"
        Case "0100": n = "4"
        Case "0101": n = "5"
        Case "0110": n = "6"
        Case "0111": n = "7"
        Case "1000": n = "8"
        Case "1001": n = "9"
        Case "1010": n = "A"
        Case "1011": n = "B"
        Case "1100": n = "C"
        Case "1101": n = "D"
        Case "1110": n = "E"
        Case "1111": n = "F"
        End Select
        nil = n & nil
    Next
    BinToHex = nil
    Err.Clear
End Function
Private Function HexToBin(ByVal Biner As String) As String
    On Error Resume Next
    ' hex to bin, used in pdu mode
    Dim bin As String
    Dim n As String
    Dim nil As String
    Dim i As Integer
    bin = ""
    For i = 1 To Len(Biner)
        bin = Mid(Biner, i, 1)
        Select Case bin
        Case "0": n = "0000"
        Case "1": n = "0001"
        Case "2": n = "0010"
        Case "3": n = "0011"
        Case "4": n = "0100"
        Case "5": n = "0101"
        Case "6": n = "0110"
        Case "7": n = "0111"
        Case "8": n = "1000"
        Case "9": n = "1001"
        Case "A": n = "1010"
        Case "B": n = "1011"
        Case "C": n = "1100"
        Case "D": n = "1101"
        Case "E": n = "1110"
        Case "F": n = "1111"
        End Select
        nil = nil & n
    Next
    HexToBin = nil
    Err.Clear
End Function
Private Function ConvToChar(ByVal hx As String) As String
    On Error Resume Next
    ' convert to char, used in pdu mode
    Dim i As Integer
    Dim tx As String
    For i = 1 To Len(hx) Step 2
        tx = tx & Chr(HexToDec(Mid(hx, i, 2)))
    Next
    ConvToChar = tx
    Err.Clear
End Function
Private Function HexToDec(ByVal x As String) As Integer
    On Error Resume Next
    ' hex to decimal, used in pdu mode
    Dim m As String
    Dim i As Byte
    Dim nil As Integer
    Dim n As Integer
    For i = 1 To 2
        m = Mid(x, i, 1)
        Select Case UCase(m)
        Case "A": n = 10
        Case "B": n = 11
        Case "C": n = 12
        Case "D": n = 13
        Case "E": n = 14
        Case "F": n = 15
        Case Else: n = CInt(m)
        End Select
        If i = 1 Then
            nil = n * 16
        Else
            nil = nil + n
        End If
    Next
    HexToDec = nil
    Err.Clear
End Function
Private Function DecToHex(ByVal x As Integer) As String
    On Error Resume Next
    ' decimal to hex, used in pdu mode
    Dim nil As String
    nil = Hex(x)
    If Len(nil) = 1 Then
        nil = "0" & nil
    End If
    DecToHex = nil
    Err.Clear
End Function
Private Function RevNum(ByVal numb As String) As String
    On Error Resume Next
    ' reverse a number, used in pdu mode
    Dim s As Integer
    Dim ma As String
    Dim b As String
    Dim a As String
    Dim ta As String
    s = 1
    ma = ""
    While (s <= Len(numb))
        ta = Mid(numb, s, 2)
        a = Mid(ta, 1, 1)
        b = Mid(ta, 2, 1)
        If b = "" Then b = "F"
        ma = ma & b & a
        s = s + 2
    Wend
    RevNum = ma
    Err.Clear
End Function
Public Property Get vUDL() As String
    On Error Resume Next
    vUDL = mvarUDL
    Err.Clear
End Property
Public Property Get vnoSCA() As String
    On Error Resume Next
    vnoSCA = mvarnoSCA
    Err.Clear
End Property
Public Property Get IndexSend() As String
    On Error Resume Next
    IndexSend = mvarIndexSend
    Err.Clear
End Property
Public Property Get vnoOA() As String
    On Error Resume Next
    vnoOA = mvarnoOA
    Err.Clear
End Property
Public Property Get vFO() As String
    On Error Resume Next
    vFO = mvarFO
    Err.Clear
End Property
Public Property Get vDCS() As String
    On Error Resume Next
    vDCS = mvarDCS
    Err.Clear
End Property
Public Property Get vSCTS_Tgl() As String
    On Error Resume Next
    vSCTS_Tgl = mvarSCTS_Tgl
    Err.Clear
End Property
Public Property Get vSCTS_Jam() As String
    On Error Resume Next
    vSCTS_Jam = mvarSCTS_Jam
    Err.Clear
End Property
Public Property Get vSCTS_Tgl_A() As String
    On Error Resume Next
    vSCTS_Tgl_A = mvarSCTS_Tgl_A
    Err.Clear
End Property
Public Property Get vSCTS_Jam_A() As String
    On Error Resume Next
    vSCTS_Jam_A = mvarSCTS_Jam_A
    Err.Clear
End Property
Public Function SMS_DecodePDU(ByVal msg As String) As String
    On Error Resume Next
    ' function to convert pdu received text to readable format
    Dim FO As String
    Dim PID As String
    Dim DCS As String
    Dim SCTS As String
    Dim UDL As String
    Dim UD As String
    Dim SCTS_Tgl As String
    Dim SCTS_Jam As String
    Dim lnSCA As String
    Dim typeSCA As String
    Dim noSCA As String
    Dim newMsg As String
    Dim lnOA As String
    Dim typeOA As String
    Dim noOA As String
    Dim SCTS_a As String
    Dim SCTS_Tgl_a As String
    Dim SCTS_Jam_a As String
    newMsg = msg
    lnSCA = HexToDec(Left(msg, 2)) * 2  'length of SCA
    newMsg = Right(newMsg, Len(newMsg) - 2)
    typeSCA = Left(newMsg, 2) '91:int,81:local
    newMsg = Right(newMsg, Len(newMsg) - 2)
    noSCA = RevNum(Left(newMsg, lnSCA - 2)) 'service center
    If UCase(Right(noSCA, 1)) = "F" Then
        noSCA = Left(noSCA, Len(noSCA) - 1)
    End If
    newMsg = Right(newMsg, Len(newMsg) - lnSCA + 2)
    FO = Left(newMsg, 2)
    newMsg = Right(newMsg, Len(newMsg) - 2)
    If FO = "06" Then
        'code of send report
        mvarIndexSend = HexToDec(Left(newMsg, 2))
        newMsg = Right(newMsg, Len(newMsg) - 2)
    End If
    'Origine Address
    lnOA = HexToDec(Left(newMsg, 2))
    If lnOA Mod 2 <> 0 Then
        lnOA = lnOA + 1
    End If
    newMsg = Right(newMsg, Len(newMsg) - 2)
    typeOA = Left(newMsg, 2)
    newMsg = Right(newMsg, Len(newMsg) - 2)
    noOA = Left(newMsg, lnOA)
    If typeOA = "D0" Then
        noOA = CharHex(noOA, 7)
    Else
        noOA = RevNum(noOA)
        If UCase(Right(noOA, 1)) = "F" Then
            noOA = Left(noOA, Len(noOA) - 1)
        End If
    End If
    newMsg = Right(newMsg, Len(newMsg) - lnOA)
    If FO <> "06" Then
        'if not report message
        PID = Left(newMsg, 2)
        newMsg = Right(newMsg, Len(newMsg) - 2)
        DCS = Left(newMsg, 2)
        newMsg = Right(newMsg, Len(newMsg) - 2)
        SCTS = RevNum(Left(newMsg, 14))
        SCTS_Tgl = Mid(SCTS, 3, 2) & "/" & Mid(SCTS, 5, 2) & "/20" & Mid(SCTS, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam = Mid(SCTS, 7, 2) & ":" & Mid(SCTS, 9, 2) & ":" & Mid(SCTS, 11, 2) 'hh:mm:dd
        newMsg = Right(newMsg, Len(newMsg) - 14)
        UDL = CInt(HexToDec(Left(newMsg, 2)))
        newMsg = Right(newMsg, Len(newMsg) - 2)
        UD = CharHex(newMsg, 7)
        UD = Left(UD, UDL)
    Else
        SCTS = RevNum(Left(newMsg, 14))
        SCTS_Tgl = Mid(SCTS, 3, 2) & "/" & Mid(SCTS, 5, 2) & "/20" & Mid(SCTS, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam = Mid(SCTS, 7, 2) & ":" & Mid(SCTS, 9, 2) & ":" & Mid(SCTS, 11, 2) 'hh:mm:dd
        newMsg = Right(newMsg, Len(newMsg) - 14)
        SCTS_a = RevNum(Left(newMsg, 14))
        SCTS_Tgl_a = Mid(SCTS_a, 3, 2) & "/" & Mid(SCTS_a, 5, 2) & "/20" & Mid(SCTS_a, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam_a = Mid(SCTS_a, 7, 2) & ":" & Mid(SCTS_a, 9, 2) & ":" & Mid(SCTS_a, 11, 2) 'hh:mm:dd
    End If
    SMS_DecodePDU = UD
    mvarnoSCA = noSCA
    mvarnoOA = noOA
    mvarFO = FO
    mvarDCS = DCS
    mvarSCTS_Tgl = SCTS_Tgl
    mvarSCTS_Jam = SCTS_Jam
    mvarSCTS_Tgl_A = SCTS_Tgl_a
    mvarSCTS_Jam_A = SCTS_Jam_a
    mvarUDL = UDL
    Err.Clear
End Function
Public Function SMS_SendPDU(ByVal DestinationNo As String, ByVal Message As String, Optional FlashMessageType As PDUFlashMessageTypeEnum = NonFlash, Optional PDUDisplayReportSMS As PDUDisplayReportSMSEnum = NoReport, Optional PDULimitPeriodOfDelivery As PDULimitPeriodOfDeliveryEnum = OneDay)
    On Error Resume Next
    ' send a sms in pdu format
    Dim SCA As String
    Dim PDU As String
    Dim MR As String
    Dim DA As String
    Dim PID As String
    Dim DCS As String
    Dim VP As String
    Dim UDL As String
    Dim UD As String
    Dim mResult As String
    Dim mAnswer As String
    Dim lenAll As Long
    Select Case FlashMessageType
    Case NonFlash
        ' normal message
        DCS = "00"
    Case Flash
        ' when received, will display to the user
        'and will be deleted from inbox of user
        DCS = "F0"
    End Select
    Select Case PDUDisplayReportSMS
    Case YesReport
        PDU = "31"
    Case NoReport
        PDU = "11"
    End Select
    Select Case PDULimitPeriodOfDelivery
    Case OneHour
        VP = "0B"
    Case TwelveHours
        VP = "8F"
    Case OneDay
        VP = "A7"
    Case TwoDays
        VP = "A8"
    Case OneWeek
        VP = "AD"
    End Select
    SCA = "00"
    MR = "00"
    'DA: Destination Address
    DA = DecToHex(Len(DestinationNo)) 'convert DestinationNo to Hex
    DA = DA & "91" '"91":Int. Number(62...),"81":Loc. Number(081..)
    DA = DA & RevNum(DestinationNo)
    PID = "00"
    UDL = DecToHex(Len(Message)) ' length of message in Hex
    UD = CharHex(Message, 8) 'Message in Hex 8bit /octet
    'Format of SMS Submit PDU
    mResult = SCA & PDU & MR & DA & PID & DCS & VP & UDL & UD
    ' find the length to send to the modem
    lenAll = Len(mResult) / 2 - 1
    mAnswer = Request("AT+CMGS=" & lenAll, , resultPos, Chr$(13))
    Select Case mAnswer
    Case ">"
        mAnswer = Request(mResult, , 4, Chr$(26))
    Case Else
        mAnswer = "ERROR"
    End Select
    SMS_SendPDU = mAnswer
    RaiseEvent Response(SMS_SendPDU)
    Err.Clear
End Function
Public Function DisConnect() As Boolean
    On Error Resume Next
    ' disconnect
    If MSComm.PortOpen = True Then MSComm.PortOpen = False
    DisConnect = MSComm.PortOpen
    Err.Clear
End Function
Public Sub ClearLog()
    On Error Resume Next
    If FileExists(LogFile) = True Then Kill LogFile
    Err.Clear
End Sub
Public Function PhoneBook_ImportOk(strFile As String) As Boolean
    On Error Resume Next
    Dim strFileData As String
    PhoneBook_ImportOk = False
    strFileData = FileData(strFile)
    strFileData = MvField(strFileData, 1, vbNewLine)
    strFileData = LCase$(strFileData)
    If strFileData = "index,cellphone no,full name" Then PhoneBook_ImportOk = True
    Err.Clear
End Function
Public Sub SMS_Export(progBar As Variant, exportFile As String)
    On Error Resume Next
    ' export all messages in the message store
    Dim rsCnt As Long
    Dim phEntry As String
    Dim nCollection As New Collection
    Set nCollection = SMS_ReadMessages
    progBar.Max = nCollection.Count
    progBar.Min = 0
    progBar.Value = 0
    FileUpdate exportFile, "Msg ID" & FM & "Message Type" & FM & "Cellphone No" & FM & "Time" & FM & "Contents", "w"
    ' loop through the messages starting from location 1 to the full capacity of the phone
    For rsCnt = 1 To nCollection.Count
        progBar.Value = rsCnt
        ' read entry at specified index
        phEntry = nCollection(rsCnt)
        ' if successfull return msg index, type,cellnumber,date,message
        If Len(phEntry) > 0 Then FileUpdate exportFile, phEntry, "a"
        DoEvents
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Public Function Phonebook_Format(progBar As Variant, PhoneBookMemoryToFormat As PhoneBookMemoryToFormatEnum) As Long
    On Error Resume Next
    Dim rsCnt As Long
    Dim pEnd As Long
    Dim delTot As Long
    Dim mResult As String
    Select Case PhoneBookMemoryToFormat
    Case SimPhoneBookFormat
        If PhoneBook_Memory <> "SM" Then Call PhoneBook_MemoryStorage(SimPhoneBook)
    Case MobileEquipmentPhoneBookFormat
        If PhoneBook_Memory <> "ME" Then Call PhoneBook_MemoryStorage(MobileEquipmentPhoneBook)
    End Select
    If PhoneBook_Capacity = 0 Then Exit Function
    delTot = 0
    pEnd = Me.PhoneBook_Capacity
    progBar.Max = pEnd
    progBar.Min = 0
    progBar.Value = 0
    For rsCnt = 1 To pEnd
        progBar.Value = rsCnt
        mResult = Me.PhoneBook_DeleteEntry(rsCnt)
        If mResult = "OK" Then delTot = delTot + 1
        DoEvents
    Next
    Phonebook_Format = delTot
    PhoneBook_MemoryStorage (ReadPhoneBookSetting)
    Err.Clear
End Function
Public Function SMS_Write(sNumber As String, sMessage As String, Optional WriteLocation As WriteLocationEnum = WriteStoUnsent) As String
    On Error Resume Next
    ' write a message on specified location and return location
    Dim mAnswer As String
    Dim sWhatToWrite As String
    Dim sNumberType As String
    Dim sLocation As String
    Dim sCommand As String
    If Left$(sNumber, 1) = "+" Then
        sNumberType = "145"
    Else
        sNumberType = "129"
    End If
    Select Case WriteLocation
    Case WriteRecRead
        sLocation = "REC READ"
    Case WriteRecUnread
        sLocation = "REC UNREAD"
    Case WriteStoSent
        sLocation = "STO SENT"
    Case WriteStoUnsent
        sLocation = "STO UNSENT"
    End Select
    sCommand = "AT+CMGW=" & Quote & sNumber & Quote & "," & sNumberType & "," & Quote & sLocation & Quote
    mAnswer = Request(sCommand, , resultPos, vbCr)
    Select Case mAnswer
    Case ">"
        mAnswer = Request(sMessage, , , Chr$(26))
        mAnswer = Replace$(mAnswer, "ýýOKý", "")
        mAnswer = MvField(mAnswer, -1, VM)
        mAnswer = Trim$(Replace$(mAnswer, "+CMGW:", ""))
    Case Else
        mAnswer = "ERROR"
    End Select
    SMS_Write = mAnswer
    RaiseEvent Response(SMS_Write)
    Err.Clear
End Function
Public Property Get OperatorName() As String
    On Error Resume Next
    Dim mResult As String
    ' get the international mobile subscriber identity
    mResult = Request("AT+COPN", , resultPos, vbCrLf)
    mResult = Trim$(Replace$(mResult, "+COPN:", ""))
    OperatorName = Replace$(mResult, Quote, "")
    RaiseEvent Response(OperatorName)
    Err.Clear
End Property
Public Function SMS_SendLocation(sLocation As String) As String
    On Error Resume Next
    ' send an sms stored on location
    Dim mAnswer As String
    mAnswer = Request("at+cmss=" & sLocation, , 4)
    SMS_SendLocation = mAnswer
    RaiseEvent Response(SMS_SendLocation)
    Err.Clear
End Function
Public Property Get NokiaTest() As String
    On Error Resume Next
    NokiaTest = Request("AT*NOKIATEST", , resultPos, vbCrLf)
    RaiseEvent Response(NokiaTest)
    Err.Clear
End Property
Public Property Get CompleteCapabilitiesList() As String
    On Error Resume Next
    Dim mResult As String
    mResult = Request("AT+GCAP", , resultPos, vbCrLf)
    mResult = Replace$(mResult, "+GCAP:", "")
    CompleteCapabilitiesList = Trim$(mResult)
    RaiseEvent Response(CompleteCapabilitiesList)
    Err.Clear
End Property
Public Property Get ActiveConfiguation() As String
    On Error Resume Next
    Dim mResult As String
    mResult = Request("AT&V", , resultPos, vbCrLf)
    ActiveConfiguation = Trim$(mResult)
    RaiseEvent Response(ActiveConfiguation)
    Err.Clear
End Property
Public Property Get Clock() As String
    On Error Resume Next
    Dim mResult As String
    mResult = Request("AT+CCLK?", , resultPos, vbCrLf)
    Clock = Trim$(mResult)
    RaiseEvent Response(Clock)
    Err.Clear
End Property

Public Property Get Connected() As Boolean
    On Error Resume Next
    Connected = Me.PortOpen
    Err.Clear
End Property

Public Property Get CurrentOperator() As String
    On Error Resume Next
    Dim mResult As String
    mResult = Request("AT+COPS?", , resultPos, vbCrLf)
    mResult = MvField(mResult, 3, ",")
    mResult = Replace$(mResult, Quote, "")
    CurrentOperator = Trim$(mResult)
    RaiseEvent Response(CurrentOperator)
    Err.Clear
End Property
Private Function DeviceGet(DeviceName As String) As Variant
    On Error Resume Next
    ' In this function we will get the devices referring to the given class name
    Dim DeviceSet As SWbemObjectSet
    Dim Device As SWbemObject
    Dim sTemp As String
    ' Set the SWbemObjectSet object
    Set DeviceSet = GetObject("winmgmts:").InstancesOf(DeviceName)
    ' Get the devices captions
    For Each Device In DeviceSet
        sTemp = sTemp & Device.Caption & "|"
    Next
    ' Remove the '|' character at the end of the string
    If Right$(sTemp, 1) = "|" Then sTemp = Left$(sTemp, Len(sTemp) - 1)
    ' Return an array (variant) with the devices captions
    DeviceGet = Split(sTemp, "|")
    Err.Clear
End Function
Public Sub DeviceProperties(ByVal sDevice As String, lstReport As Variant)
    On Error Resume Next
    ' This function returns all the properties of a specific device and loads them to the listview
    Dim DeviceSet As SWbemObjectSet
    Dim Device As SWbemObject
    Dim vTemp As Variant
    Dim sTemp As String
    Dim mProperties As Variant
    Dim vItems As Variant
    lstReport.ListItems.Clear
    lstReport.View = 3
    lstReport.Checkboxes = False
    lstReport.GridLines = True
    lstReport.ColumnHeaders.Clear
    lstReport.ColumnHeaders.Add , , "Property"
    lstReport.ColumnHeaders.Add , , "Value"
    ' Set theSWbemObjectSet object
    Set DeviceSet = GetObject("winmgmts:").InstancesOf("Win32_POTSModem")
    For Each Device In DeviceSet
        ' Check if the current device in the chosen device
        If LCase$(Device.Caption) = LCase$(sDevice) Then
            ' Get all the properties of the chosen device
            For Each vTemp In Device.Properties_
                If vTemp <> "" And vTemp <> vbNull Then
                    ' Add the property name and its value to the temporary string
                    sTemp = sTemp & vTemp.Name & "^" & vTemp & "|"
                End If
            Next
            ' Remove the '|' character at the end of the string
            If Right$(sTemp, 1) = "|" Then
                sTemp = Left$(sTemp, Len(sTemp) - 1)
            End If
        End If
    Next
    ' Return an array containing the device properties
    mProperties = Split(sTemp, "|")
    ' Populate the ListView with the device's properties
    For Each vTemp In mProperties
        vItems = Split(vTemp, "^")
        lstReport.ListItems.Add(, , CStr(vItems(0))).SubItems(1) = vItems(1)
    Next
    LstViewAutoResize lstReport
    Err.Clear
End Sub
Public Function Devices() As Collection
    On Error Resume Next
    ' get the names of modems
    Dim DevicesNames As Variant
    Dim Device As Variant
    Dim TempDevice As Variant
    Dim NumOfDevices As Integer
    Dim theDevices As Variant
    Dim tmpDevice As String
    Set Devices = New Collection
    ' This array contains the Computer System Hardware Classes names, we will only look at the modems
    DevicesNames = Array("Win32_POTSModem")
    ', "Win32_POTSModemToSerialPort"
    ' Find the number of hardware classes
    NumOfDevices = UBound(DevicesNames)
    ' Find all the hardware devices
    For Each Device In DevicesNames
        ' Make sure that the operating system can process other events
        DoEvents
        tmpDevice = Right$(Device, Len(Device) - 6)
        theDevices = DeviceGet(CStr(Device))
        For Each TempDevice In theDevices
            Devices.Add CStr(TempDevice)
        Next
    Next
    Err.Clear
End Function
Public Sub PhoneBook_ListView(progBar As Variant, LstView As Variant, used As Long, capacity As Long, Optional mIcon As String = "", Optional mSmallIcon As String = "")
    On Error Resume Next
    ' load contents of the selected phonebook
    ' to the listview
    Dim rsCnt As Long
    Dim phEntry As String
    Dim lstItem As Variant
    Dim spLine() As String
    Dim lstTotal As Long
    progBar.Max = capacity
    progBar.Min = 0
    progBar.Value = 0
    ' create headings
    LstViewMakeHeadings LstView, "Index,Cellphone No,Full Name"
    ' loop through the phonebook starting from location 1 to the full capacity of the
    ' phone
    For rsCnt = 1 To capacity
        progBar.Value = rsCnt
        ' read entry at specified index
        phEntry = Me.PhoneBook_ReadEntry(rsCnt)
        ' if successfull return index,cellno,fullname
        If Len(phEntry) > 0 Then
            spLine = Split(phEntry, ",")
            Set lstItem = LstView.ListItems.Add(, , spLine(0))
            lstItem.SubItems(1) = spLine(1)
            lstItem.SubItems(2) = spLine(2)
            If Len(mIcon) > 0 Then lstItem.Icon = mIcon
            If Len(mSmallIcon) > 0 Then lstItem.SmallIcon = mSmallIcon
        End If
        DoEvents
        ' ensure that if we have reached the used limit, exit the loop
        ' we do not want to read the empty contacts anymore
        lstTotal = LstView.ListItems.Count
        If lstTotal = used Then Exit For
    Next
    progBar.Value = 0
    Err.Clear
End Sub
Private Sub LstViewMakeHeadings(LstView As Variant, ByVal strHeads As String)
    On Error Resume Next
    ' used to create columns in a listview
    Dim fldCnt As Integer
    Dim FldHead() As String
    Dim fldTot As Integer
    Dim colX As Variant
    FldHead = Split(strHeads, ",")
    fldTot = UBound(FldHead)
    LstView.ColumnHeaders.Clear
    LstView.ListItems.Clear
    LstView.Sorted = False
    ' first column should be left aligned
    Set colX = LstView.ColumnHeaders.Add(, , FldHead(0), 1440)
    For fldCnt = 1 To fldTot
        Set colX = LstView.ColumnHeaders.Add(, , FldHead(fldCnt), 1440)
    Next
    LstView.View = 3
    LstView.Checkboxes = True
    LstView.GridLines = True
    LstView.FullRowSelect = True
    LstView.Refresh
    Err.Clear
End Sub
Public Function FileToken(ByVal strFileName As String, Optional ByVal Sretrieve As String = "F", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    ' get a file token like path, directory, extension etc
    Dim intNum As Long
    Dim sNew As String
    FileToken = strFileName
    Select Case UCase$(Sretrieve)
    Case "D"
        FileToken = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            FileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            FileToken = Mid$(strFileName, 1, intNum - 1)
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            FileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        FileToken = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            FileToken = Left$(strFileName, intNum - 1)
        End If
    End Select
    Err.Clear
End Function
Sub CreateNestedDirectory(ByVal StrCompletePath As String)
    On Error Resume Next
    ' create a nested directory
    Dim spPaths() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim curPath As String
    Call StrParse(spPaths, StrCompletePath, "\")
    spTot = UBound(spPaths)
    For spCnt = 1 To spTot
        curPath = MvFromMv(StrCompletePath, 1, spCnt, "\")
        If DirExists(curPath) = False Then
            MkDir curPath
        End If
    Next
    Err.Clear
End Sub
Public Function StrParse(retarray() As String, ByVal strText As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    ' works like the split function, but starting at 1
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    If Len(Delim) = 0 Then Delim = VM
    varArray = Split(strText, Delim)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
    Next
    StrParse = UBound(retarray)
    Err.Clear
End Function
Public Function MvFromMv(ByVal strOriginalMv As String, ByVal startPos As Long, Optional ByVal NumOfItems As Long = -1, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    ' create a delimited string from another
    Dim sporiginal() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sLine As String
    Dim endPos As Long
    sLine = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StrParse(sporiginal, strOriginalMv, Delim)
    spTot = UBound(sporiginal)
    If NumOfItems = -1 Then
        endPos = spTot
    Else
        endPos = (startPos + NumOfItems) - 1
    End If
    For spCnt = startPos To endPos
        If spCnt = endPos Then
            sLine = sLine & sporiginal(spCnt)
        Else
            sLine = sLine & sporiginal(spCnt) & Delim
        End If
    Next
    MvFromMv = sLine
    Err.Clear
End Function
Private Function DirExists(ByVal Sdirname As String) As Boolean
    On Error Resume Next
    ' does the directory exist
    Dim sDir As String
    DirExists = False
    sDir = Dir$(Sdirname, vbDirectory)
    If Len(sDir) > 0 Then DirExists = True
    Err.Clear
End Function
Public Sub LoadDevices(cboBox As Variant)
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim myDevices As New Collection
    Set myDevices = Me.Devices
    cboBox.Clear
    rsTot = myDevices.Count
    For rsCnt = 1 To rsTot
        cboBox.AddItem myDevices(rsCnt)
    Next
    Err.Clear
End Sub
Private Sub LstViewAutoResize(LstView As Variant)
    On Error Resume Next
    Dim col2adjust As Long
    Dim col2adjust_Tot As Long
    If LstView.ListItems.Count = 0 Then
        Err.Clear
        Exit Sub
    End If
    col2adjust_Tot = LstView.ColumnHeaders.Count - 1
    For col2adjust = 0 To col2adjust_Tot
        Call SendMessage(LstView.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
    'LstViewResizeMax lstView
    Err.Clear
End Sub
Public Function LstViewFindItem(LstView As Variant, ByVal StrSearch As String, Optional ByVal SearchWhere As FindWhere = Search_Text, Optional SearchItemType As SearchType = Search_Whole) As Long
    On Error Resume Next
    Dim itmFound As ListItem
    LstViewFindItem = 0
    Set itmFound = LstView.FindItem(StrSearch, SearchWhere, , SearchItemType)
    If TypeName(itmFound) = "Nothing" Then
        Err.Clear
        Exit Function
    End If
    LstViewFindItem = CLng(itmFound.Index)
    Set itmFound = Nothing
    Err.Clear
End Function
Public Function LstViewGetRow(LstView As Variant, ByVal idx As Long) As Variant
    On Error Resume Next
    Dim retarray() As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    clsColTot = LstView.ColumnHeaders.Count
    ReDim retarray(clsColTot)
    retarray(1) = LstView.ListItems(idx).Text
    clsColTot = clsColTot - 1
    For clsColCnt = 1 To clsColTot
        retarray(clsColCnt + 1) = LstView.ListItems(idx).SubItems(clsColCnt)
    Next
    LstViewGetRow = retarray
    Err.Clear
End Function
Public Function ExtractNumbers(ByVal strValue As String) As String
    On Error Resume Next
    Dim i As Long
    Dim sResult As String
    Dim iLen As Long
    Dim myStr As String
    sResult = ""
    iLen = Len(strValue)
    For i = 1 To iLen
        myStr = Mid$(strValue, i, 1)
        If InStr("-0123456789", myStr) > 0 Then
            sResult = sResult & myStr
        End If
    Next
    ExtractNumbers = sResult
    Err.Clear
End Function
