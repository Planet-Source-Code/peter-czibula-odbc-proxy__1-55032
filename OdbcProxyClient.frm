VERSION 5.00
Begin VB.Form frmTestODBCproxy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Odbc Proxy Demo"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDSN 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "OdbcProxyClient.frx":0000
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton cmdNEXT 
      Caption         =   "Movenext"
      Height          =   350
      Left            =   1800
      TabIndex        =   14
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdODBC 
      Caption         =   "Connect to DSN"
      Height          =   350
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtInput 
      Height          =   5655
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   350
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   555
      Width           =   1575
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   555
      Width           =   1575
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "Execute SQL"
      Height          =   350
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtSQL 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "OdbcProxyClient.frx":003A
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "3000"
      Top             =   105
      Width           =   615
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "192.9.120.8"
      Top             =   105
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Enter Remote DSN here:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Data Received from ODBC proxy Server:"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Enter SQL statement here:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Not Connected"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmTestODBCproxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstWS1 As New rstWinsock

Private Sub Form_Load()
    txtInput.Text = ""
    cmdDisconnect.Enabled = False
    cmdODBC.Enabled = False
    cmdNEXT.Enabled = False
    cmdSQL.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'disconnect if the form is closed
    rstWS1.CloseSocket
End Sub

Private Sub cmdConnect_Click()
    'connect to specified server at the specified port
    Set rstWS1 = New rstWinsock
    rstWS1.RemoteHostPort = txtPort.Text
    rstWS1.RemoteHostIP = txtServer.Text
    rstWS1.ConnectSocket
    If rstWS1.State <> Connected Then
        MsgBox "Error while connecting: " & rstWS1.StateDescr
    Else
        Label3.Caption = "Connected to " + rstWS1.RemoteHostIP
        cmdConnect.Enabled = False
        cmdDisconnect.Enabled = True
        cmdODBC.Enabled = True
        cmdNEXT.Enabled = False
        cmdSQL.Enabled = False
    End If
End Sub

Private Sub cmdDisconnect_Click()
    'close the connection
    rstWS1.CloseSocket
    If rstWS1.State <> notConnected Then
        MsgBox "Error while deconnecting: " & rstWS1.StateDescr
    Else
        Label3.Caption = "Not Connected"
        cmdDisconnect.Enabled = False
        cmdConnect.Enabled = True
        cmdODBC.Enabled = False
        cmdNEXT.Enabled = False
        cmdSQL.Enabled = False
    End If
End Sub

Private Sub cmdODBC_Click()
    Label3.Caption = "Sending ODBC DSN"
    rstWS1.ConnectRemoteDSN txtDSN
    If rstWS1.State <> Connected Then
        txtInput.Text = "Error while connecting to DSN: " & rstWS1.ErrDescr
    Else
        txtInput.Text = "Connected to DSN " + rstWS1.RemoteHostIP
        cmdNEXT.Enabled = True
        cmdSQL.Enabled = True
    End If
    Label3.Caption = "ODBC DSN Sent"
End Sub

Private Sub cmdSQL_Click()
    Dim t As Long
    Label3.Caption = "Sending SQL"
    rstWS1.ExecuteSQL txtSQL.Text
    If rstWS1.State <> Connected Then
        txtInput.Text = "Error while executing SQL: " & rstWS1.ErrDescr
    Else
        txtInput.Text = rstWS1.Fields.Count & " Fields." & vbCrLf
        For t = 0 To rstWS1.Fields.Count - 1
            txtInput.Text = txtInput.Text + rstWS1.Fields(t).name & "=" & rstWS1.Fields(t).Value & vbCrLf
        Next t
    End If
    Label3.Caption = "SQL Sent"
End Sub

Private Sub cmdNEXT_Click()
    Dim t As Long
    Label3.Caption = "Sending MOVENEXT"
    rstWS1.MoveNext
    If rstWS1.State <> Connected Then
        txtInput.Text = "Error while executing MOVENEXT: " & rstWS1.ErrDescr
    Else
        If rstWS1.EOF Then
            txtInput.Text = "EOF reached."
        Else
            txtInput.Text = rstWS1.Fields.Count & " Fields." & vbCrLf
            For t = 0 To rstWS1.Fields.Count - 1
                txtInput.Text = txtInput.Text + rstWS1.Fields(t).name & "=" & rstWS1.Fields(t).Value & vbCrLf
            Next t
        End If
    End If
    Label3.Caption = "MOVENEXT Sent"
End Sub
