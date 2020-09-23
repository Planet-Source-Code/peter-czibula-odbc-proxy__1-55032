VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rstWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************
'*
'* Make sure you add the following REFERENCES to your project
'* = Microsoft ActiveX Data Objects 2.7 Library
'*          Reference=C:\Program Files\Common Files\system\ado\msado15.dll
'*
'* Relies on CSocket class - Copyright © 2002 by Oleg Gdalevich
'* Visual Basic Internet Programming website (http://www.vbip.com)
'* which also needs MSocketSupport code module
'* Permits the use of Winsock functionality in a class
'**************************************

Option Explicit

'***************************************
'* Exposed Class Properties
'***************************************

Public RemoteHostIP As String
Public RemoteHostPort As Long
Public State As rstWinsockState
Public Enum rstWinsockState
    Connected = 0
    notConnected = -1
    TimeOut = 1
    Error = 4
End Enum
Public Response As rstWinsockResponse
Public Enum rstWinsockResponse
    Waiting = 0
    Received = 1
End Enum
Public TimeoutSecs As Long
Public ErrDescr As String
Public ErrCode As Long
Public Fields As ADODB.Fields
Public EOF As Boolean

'***************************************
'* Class variables
'***************************************

'Private objSocket As New CSocket
Private strDataReceived As String
Private dsData As wsDataStream
Private Type wsDataStream
    StxEtx As Byte
    LenHi As Byte
    LenLo As Byte
    Unused1 As Byte
    Unused2 As Byte
    Cmd As Byte
    Data As String
End Type
Private rst As ADODB.Recordset

'***************************************
'* WINSOCK DECLARATIONS
'***************************************

Private WithEvents objSocket As CSocket
Attribute objSocket.VB_VarHelpID = -1

Const sckClosed = 0             'Default. Closed
Const sckOpen = 1               'Open
Const sckListening = 2          'Listening
Const sckConnectionPending = 3  'Connection pending
Const sckResolvingHost = 4      'Resolving host
Const sckHostResolved = 5       'Host resolved
Const sckConnecting = 6         'Connecting
Const sckConnected = 7          'Connected
Const sckClosing = 8            'Peer is closing the connection
Const sckError = 9              'Error

'***************************************
'* DBTCP CLIENT DECLARATIONS
'***************************************

Const CHR_NULL = &H0
Const CHR_STX = &H5
Const CHR_ETX = &H6
Const CHR_CR = &HD
Const CHR_LF = &HA
Const OK = 0
Const Err = -1

Const FETCH_OK = 0
Const FETCH_ERROR = -1
Const FETCH_EOF = 1

' Client command code
Const CLN_CMD_CONNECT = 1        ' Connect
Const CLN_CMD_SQL = 2            ' Execute SQL
Const CLN_CMD_FETCH_F = 5        ' Fetch result set fields info
Const CLN_CMD_FETCH = 6          ' Fetch next row in result set

' Server return code
Const SRV_CMD_ERROR = 50         ' Generic Error
Const SRV_CMD_SUCCESS = 51       ' Success
Const SRV_FETCH_EOF = 52         ' Fetch: end of records
Const SRV_CMD_UNDEFINED = 53     ' Server: command id undefined
Const SRV_FETCH_RESULT = 54      ' Fetch: packet is a data record
Const SRV_FETCH_FIELDS = 55      ' Fetch: packet is fields description
Const SRV_CMD_NO_DATA = 56       ' Sql: Query has no result

'***************************************
'* ADO RECORDSET DECLARATIONS
'***************************************

' MS Access..............SQL Server........ADO Constant
' ------------------------------------------------------
' Yes/No.................Bit...............adBoolean
' Number(Byte)...........TinyInt...........adTinyInt
' Number(Integer)........SmallInt..........adSmallInt
' Number(Long Integer)...Int...............adInteger
' Number(Single).........Real..............adSingle
' Number(Double).........Float.............adDouble
' Currency...............Money.............adCurrency
' .......................Smallmoney........adCurrency
' Decimal/Numeric........Decimal...........adDecimal
' .......................Numeric...........adNumeric
' Date/Time..............Datetime..........adDate
' .......................Smalldatetime.....adDBDate
' AutoNumber.............Int...............adInteger
' Text(n)................Varchar(n)........adVarchar
' .......................Nvarchar(n).......adVarchar
' Memo...................Text..............adLongVarWChar
' OLE Object.............Image.............adLongVarBinary
' Replication ID.........Uniqueidentifier..adGUID

Const adEmpty = 0
Const adTinyInt = 16                'Number(Byte)
Const adSmallInt = 2                'Number(Integer)
Const adInteger = 3                 'Number(Long Integer)
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14                'Decimal/Numeric
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7                    'Date/Time
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200               'Text(n)
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205
Const adChapter = 136
Const adFileTime = 64
Const adPropVariant = 138
Const adVarNumeric = 139
Const adArray = &H2000

Private Sub Class_Initialize()

'***************************************
'* On Class initialize, create the winsock object
'***************************************
    
    ' Set default Timeout
    TimeoutSecs = 10
    ClearError (notConnected)
    Set Fields = Nothing
    Set rst = Nothing
End Sub

Private Sub Class_Terminate()
    
'***************************************
'* On Class terminate, remove the winsock from memory
'***************************************
    
    On Error GoTo SocketAlreadyClosed
    If objSocket.State <> sckClosed Then CloseSocket
    Set objSocket = Nothing
SocketAlreadyClosed:

End Sub

Public Sub ConnectSocket()
    
'***************************************
'* Procedure: ConnectSocket
'***************************************
    
    Dim d As Date
    
    Set objSocket = New CSocket
    ' loop until connected or timeout
    d = Now()
    objSocket.Connect RemoteHostIP, RemoteHostPort
    Do While objSocket.State <> sckConnected
        If DateDiff("s", d, Now()) > TimeoutSecs Then
            Exit Do
        End If
        DoEvents
    Loop
    If objSocket.State = sckConnected Then
        ClearError (Connected)
    Else
        ClearError (TimeOut)
    End If
    Set Fields = Nothing
    Set rst = Nothing
End Sub

Public Sub CloseSocket()
    
'***************************************
' Procedure: CloseSocket
'***************************************
    
    Dim d As Date
    
    On Error GoTo SocketAlreadyClosed
    ' loop until response received or timeout
    d = Now()
    objSocket.CloseSocket
    Do While objSocket.State = sckConnected
        If DateDiff("s", d, Now()) > TimeoutSecs Then
            Exit Do
        End If
        DoEvents
    Loop
    If objSocket.State <> sckConnected Then
        ClearError (notConnected)
    Else
        ClearError (TimeOut)
    End If
    Set objSocket = Nothing
    Set Fields = Nothing
    Set rst = Nothing
SocketAlreadyClosed:
End Sub

Public Sub ConnectRemoteDSN(RemoteDSN As String)
    
'***************************************
'* Procedure: ConnectRemoteDSN
'***************************************
    
    Dim sHead As String
    Dim sData As String
    Dim d As Date
    
    If objSocket.State = sckConnected Then
        ' prepare the data (remove line breaks, tabs ...)
        sData = RemoteDSN
        sData = Replace(sData, vbTab, "")
        sData = Replace(sData, vbCr, "")
        sData = Replace(sData, vbLf, "")
        sData = Trim(sData)
        ' check for ; at the end
        If Right(sData, 1) <> ";" Then sData = sData & ";"
        ' prepare the Header
        sHead = Chr(CHR_STX) & Chr(Len(sData) \ 255) & Chr(Len(sData) Mod 255) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CLN_CMD_CONNECT)
        ' put Wait Flag
        Response = Waiting
        ' send it
        objSocket.SendData sHead & sData
        ' loop until response received or timeout
        d = Now()
        Do While Response = Waiting
            If DateDiff("s", d, Now()) > TimeoutSecs Then
                Exit Do
            End If
            DoEvents
        Loop
        If Response = Received Then
            ' check the response
            If dsData.Cmd = SRV_CMD_SUCCESS Then
                ClearError (Connected)
            Else
                ' we have an incorrect reponse
                SetError Val(dsData.Cmd)
            End If
        Else
            ClearError (TimeOut)
        End If
    Else
        ClearError (notConnected)
    End If
    Set Fields = Nothing
    Set rst = Nothing
End Sub

Public Sub ExecuteSQL(SQL As String)
    
'***************************************
'* Procedure: ExecuteSQL
'***************************************
    
    Dim sHead As String
    Dim sData As String
    Dim d As Date
    
    If objSocket.State = sckConnected Then
        ' prepare the data (remove line breaks, tabs ...)
        sData = SQL
        sData = Replace(sData, vbTab, "")
        sData = Replace(sData, vbCr, "")
        sData = Replace(sData, vbLf, "")
        sData = Trim(sData)
        ' check for ; at the end
        If Right(sData, 1) <> ";" Then sData = sData & ";"
        sData = sData & vbCrLf
        ' prepare the Header
        sHead = Chr(CHR_STX) & Chr(Len(sData) \ 255) & Chr(Len(sData) Mod 255) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CLN_CMD_SQL)
        ' put Wait Flag
        Response = Waiting
        ' send it
        objSocket.SendData sHead & sData
        ' loop until response received or timeout
        d = Now()
        Do While Response = Waiting
            If DateDiff("s", d, Now()) > TimeoutSecs Then
                Exit Do
            End If
            DoEvents
        Loop
        If Response = Received Then
            ' check the response
            If dsData.Cmd = SRV_CMD_SUCCESS Then
                ' the command was successful, get the field description from the server
                If UCase(Left(SQL, 6)) = "SELECT" Then
                    ' we asked for a SELECT statement and
                    ' the SQL statement returns a full recordset
                    ' -> retreive fields
                    GetFields
                    If Response = Received Then
                        ' the server returned the field description, load the first record
                        MoveNext
                        If State = Connected Then
                            ' the server returned the first record correctly
                            ClearError (Connected)
                            EOF = False
                        End If
                    Else
                        ' we did not receive the field description correctly
                        SetError Val(dsData.Cmd)
                    End If
                Else
                    ' we have a different command than the SELECT statement
                    ' and received it was successful
                    ClearError (Connected)
                    EOF = True
                End If
            ElseIf dsData.Cmd = SRV_CMD_NO_DATA Then
                ' the SQL statement returns an empty recordset
                ClearError (Connected)
                EOF = True
                Set Fields = Nothing
                Set rst = Nothing
            Else
                ' we have an incorrect reponse
                SetError Val(dsData.Cmd)
            End If
        Else
            ClearError (TimeOut)
        End If
    Else
        ClearError (notConnected)
    End If
End Sub

Private Sub GetFields()
    
'***************************************
'* Procedure: GetFields (private use)
'* Send a fetch field description to the odbc proxy
'* and fill in our FIELDS collection (field name, type and len)
'***************************************
    
    Dim sHead As String
    Dim rFields() As String
    Dim fName As String
    Dim d As Date, tt As Long
    Dim RetryCount As Long
    
    ' prepare the Header
    sHead = Chr(CHR_STX) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CLN_CMD_FETCH_F)
    RetryCount = 1
Retry:
    ' put Wait Flag
    Response = Waiting
    ' send it
    objSocket.SendData sHead
    ' loop until response received or timeout
    d = Now()
    Do While Response = Waiting
        If DateDiff("s", d, Now()) > TimeoutSecs Then
            Exit Do
        End If
        DoEvents
    Loop
    If Response = Received Then
        ' check the response
        If dsData.Cmd = SRV_FETCH_FIELDS Then
            ' the server returned a field description collection
            Set rst = Nothing
            Set rst = New ADODB.Recordset
            rFields = Split(dsData.Data, Chr(CHR_NULL))
            For tt = 1 To UBound(rFields) - 1 Step 2
                fName = Mid(rFields(tt), 2, Len(rFields(tt)) - 2)
                Select Case Right(rFields(tt), 1)
                    Case "N"
                        Select Case Asc(rFields(tt + 1))
                            Case 5
                                ' type is Numeric, Len = 5
                                rst.Fields.Append fName, adSmallInt
                            Case 10
                                ' type is Numeric, Len = 10
                                rst.Fields.Append fName, adInteger
                            Case 30
                                ' type is Numeric, Len = 30
                                rst.Fields.Append fName, adCurrency
                        End Select
                    Case "C"
                        rst.Fields.Append fName, adVarChar, Asc(rFields(tt + 1))
                    Case "D"
                        rst.Fields.Append fName, adDate
                End Select
            Next tt
            rst.Open
            Set Fields = Nothing
            Set Fields = rst.Fields
            rst.AddNew
            ClearError (Connected)
        Else
            ' attention return code 95 is not documented!
            ' try again (max 5 times)
            If RetryCount < 5 Then
                RetryCount = RetryCount + 1
                GoTo Retry
            End If
            ' we have an incorrect reponse
            SetError Val(dsData.Cmd)
        End If
    Else
        ClearError (TimeOut)
    End If

End Sub

Public Sub MoveNext()
    
'***************************************
'* Procedure: Movenext (public use)
'* Send a fetch next record to the odbc proxy
'* and fill in our FIELDS collection (field value)
'***************************************
    
    Dim sHead As String
    Dim rFields() As String
    Dim fValue As String
    Dim d As Date, tt As Long, pt As Long
    
    ' check if we have send a SELECT before
    ' if not the following statement will generate an VB error
    On Error GoTo NoFields
    If Fields.Count > 0 Then
        On Error GoTo 0
        ' prepare the Header
        sHead = Chr(CHR_STX) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CHR_NULL) & Chr(CLN_CMD_FETCH)
        ' put Wait Flag
        Response = Waiting
        ' send it
        objSocket.SendData sHead
        ' loop until response received or timeout
        d = Now()
        Do While Response = Waiting
            If DateDiff("s", d, Now()) > TimeoutSecs Then
                Exit Do
            End If
            DoEvents
        Loop
        If Response = Received Then
            ' check the response
            If dsData.Cmd = SRV_FETCH_RESULT Then
                ' the server returned a field collection
                pt = 1
                For tt = 0 To Fields.Count - 1
                    If Mid(dsData.Data, pt, 1) = Chr(CHR_NULL) Then
                        fValue = Mid(dsData.Data, pt + 2, Asc(Mid(dsData.Data, pt + 1, 1)))
                        Select Case Fields(tt).Type
                            Case adSmallInt
                                ' type is Numeric, Len = 5
                                Fields(tt).Value = CLng(fValue)
                            Case adInteger
                                ' type is Numeric, Len = 10
                                Fields(tt).Value = CSng(fValue)
                            Case adCurrency
                                ' type is Numeric, Len = 30
                                fValue = Replace(fValue, ",", ".")
                                Fields(tt).Value = Val(fValue)
                            Case adVarChar
                                ' type is Char, truncate to defined size
                                Fields(tt).Value = Mid(fValue, 1, Fields(tt).DefinedSize)
                            Case adDate
                                ' type is Date, national independant
                                If fValue <> "" Then
                                    ' ATTENTION: date format returned has changed with dbTCP version 2
                                    ' this is format for version 2
                                    Fields(tt).Value = CDate(#1/1/1800#)
                                    Fields(tt).Value = DateAdd("yyyy", CLng(Mid(fValue, 1, 4)) - 1800, Fields(tt).Value)
                                    Fields(tt).Value = DateAdd("m", CLng(Mid(fValue, 6, 2)) - 1, Fields(tt).Value)
                                    Fields(tt).Value = DateAdd("d", CLng(Mid(fValue, 9, 2)) - 1, Fields(tt).Value)
                                Else
                                    Fields(tt).Value = Empty
                                End If
                        End Select
                    Else
                        SetError Val(dsData.Cmd)
                        Exit For
                    End If
                    pt = pt + Asc(Mid(dsData.Data, pt + 1, 1)) + 2
                Next tt
                rst.Update
                ClearError (Connected)
                EOF = False
            ElseIf dsData.Cmd = SRV_FETCH_EOF Then
                ClearError (Connected)
                EOF = True
            Else
                ' we have an incorrect reponse
                SetError Val(dsData.Cmd)
            End If
        Else
            ClearError (TimeOut)
        End If
    End If
    Exit Sub
NoFields:
    SetError Val(dsData.Cmd)
End Sub

Private Sub objSocket_OnDataArrival(ByVal bytesTotal As Long)
    
'***************************************
'* event objSocket_DataArrival
'* Handles data arriving from the ODBC proxy.
'***************************************
    
    objSocket.GetData strDataReceived
    ' format data
    dsData.StxEtx = Asc(Mid(strDataReceived, 1, 1))
    dsData.LenHi = Asc(Mid(strDataReceived, 2, 1))
    dsData.LenLo = Asc(Mid(strDataReceived, 3, 1))
    dsData.Unused1 = Asc(Mid(strDataReceived, 4, 1))
    dsData.Unused2 = Asc(Mid(strDataReceived, 5, 1))
    dsData.Cmd = Asc(Mid(strDataReceived, 6, 1))
    dsData.Data = Mid(strDataReceived, 7)
    ' reset the flag
    Response = Received

End Sub

Private Sub objSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

'***************************************
'* event objSocket_Error
'* Handles error messages from the socket.
'* Sets our Err Object
'***************************************
    
    State = Error
    ErrCode = Number
    ErrDescr = "Winsock Error: " & Number & "-" & Description
End Sub

Private Sub SetError(code As Long)
    
'***************************************
'* Returns textual description of ODBC proxy error
'***************************************
    
    ErrCode = code
    Select Case code
        Case SRV_CMD_ERROR
            ErrDescr = "Server returned: Generic Error Code: " & code & "-" & dsData.Data
        Case SRV_CMD_SUCCESS
            ErrDescr = "Server returned: Success" & code & "-" & dsData.Data
        Case SRV_FETCH_EOF
            ErrDescr = "Server returned: Fetch: end of records" & code & "-" & dsData.Data
        Case SRV_CMD_UNDEFINED
            ErrDescr = "Server returned: Server: command id undefined" & code & "-" & dsData.Data
        Case SRV_FETCH_RESULT
            ErrDescr = "Server returned: Fetch: packet is a data record" & code & "-" & dsData.Data
        Case SRV_FETCH_FIELDS
            ErrDescr = "Server returned: Fetch: packet is fields description" & code & "-" & dsData.Data
        Case SRV_CMD_NO_DATA
            ErrDescr = "Server returned: Sql: Query has no result" & code & "-" & dsData.Data
        Case Else
            ErrDescr = "Server returned: Undefined Error Code " & code & "-" & dsData.Data
    End Select
    State = Error
    EOF = True
    Set Fields = Nothing
    Set rst = Nothing

End Sub

Private Sub ClearError(code As rstWinsockState)
    
'***************************************
'* Clears the error code variables
'***************************************
    
    State = code
    ErrCode = 0
    ErrDescr = ""

End Sub

Property Get StateDescr() As String
    
'***************************************
'* Returns textual description of the Class State
'***************************************
    
    Select Case State
        Case Connected
            StateDescr = "Connected"
        Case notConnected
            StateDescr = "Not Connected"
        Case TimeOut
            StateDescr = "Timeout"
        Case Error
            StateDescr = "Error"
    End Select
End Property

