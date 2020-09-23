VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Client"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   210
      Top             =   630
   End
   Begin VB.CommandButton cmdAddCRLF 
      Caption         =   "Add Line Feed"
      Height          =   330
      Left            =   4725
      TabIndex        =   6
      Top             =   2835
      Width           =   1275
   End
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   2835
      Width           =   4530
   End
   Begin VB.TextBox txtData 
      Height          =   2220
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "fMain.frx":0000
      Top             =   525
      Width           =   5895
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   330
      Left            =   4830
      TabIndex        =   3
      Top             =   105
      Width           =   1065
   End
   Begin VB.TextBox txtPort 
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      Text            =   "Port"
      Top             =   105
      Width           =   1065
   End
   Begin VB.TextBox txtHostName 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Text            =   "IP Address"
      Top             =   105
      Width           =   2325
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   330
      Left            =   3675
      TabIndex        =   0
      Top             =   105
      Width           =   1065
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long

Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Private Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long

Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Private Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function CloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long

Private Declare Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long

Private Declare Function Recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Private Declare Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long


Private Const SOCK_STREAM = 1    'Stream socket
Private Const AF_INET = 2        'Internetwork: UDP, TCP e.t.c
Private Const IPPROTO_TCP = 6    'TCP


Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

Private Const SCK_VERSION1 = &H101                  'Windows sockets version 1.1
Private Const SCK_VERSION2 = &H202                  'Windows sockets version 2.2

Private Const OFFSET = 65536

Private Const FIONBIO = &H8004667E

Private Type WSAData
    WVersion        As Integer                      'Version
    WHighVersion    As Integer                      'High Version
    szDescription   As String * WSADESCRIPTION_LEN  'Description
    szSystemStatus  As String * WSASYS_STATUS_LEN   'Status of system
    iMaxSockets     As Integer                      'Maximum number of sockets allowed
    iMaxUdpDg       As Integer                      'Maximum UDP datagrams
    lpVendorInfo    As Long                         'Vendor Info
End Type


'Socket Address structure
Private Type SOCKADDR_IN
    sin_family          As Integer      'Address family
    sin_port            As Integer      'Port
    sin_addr            As Long         'Long address
    sin_zero(1 To 8)    As Byte         'Not used by us
End Type


Private WSAInfo         As WSAData
Private lngSocketHandle As Long


Private Sub cmdAddCRLF_Click()
    'Add a vbcrlf to the end of the line.
    Let txtSend.Text = txtSend.Text & vbCrLf
End Sub


Private Sub cmdConnect_Click()

  Dim lngHostName As String
  Dim udtSockaddr As SOCKADDR_IN
    
    'Switch the socket so that it does not block.
    Call ioctlsocket(lngSocketHandle, FIONBIO, 1)
    
    'First create a new streat socket that uses the TCP protocol
    Let lngSocketHandle = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    'Error checking
    If lngSocketHandle <= 0 Then
        Let txtData.Text = txtData.Text & "*** ERROR: Could not create socket!" & vbCrLf
        Let txtData.SelStart = Len(txtData)
        Exit Sub
    End If

    'Try to resolve the host name to the long address.
    'This function will work if it was an IP
    Let lngHostName = inet_addr(txtHostName.Text)

    'If it didn't work...
    If lngHostName <= 0 Then
        'Try the gethostbyname function. This will work for aliases like "www.microsoft.com"
        Let lngHostName = gethostbyname(txtHostName.Text)
           
        'Some more error checking
        If lngHostName <= 0 Then
            Let txtData.Text = txtData.Text & "*** ERROR: Could not resolve host " & txtHostName.Text & "!" & vbCrLf
            Let txtData.SelStart = Len(txtData)
            Exit Sub
        End If
    End If
    
    If Not IsNumeric(txtPort.Text) Then
        Let txtData.Text = txtData.Text & "*** ERROR: Invalid port (Must be numeric)!" & vbCrLf
        Let txtData.SelStart = Len(txtData)
        Exit Sub
    End If
    
    'Fill out the socket address structure to ready the connection
    With udtSockaddr
        Let .sin_family = AF_INET
        Let .sin_addr = lngHostName
        Let .sin_port = htons(CInt(txtPort.Text))   'Remember: Network byte order
    End With
    
    'Switch the socket so that it blocks and the connect function returns straight away
    Call ioctlsocket(lngSocketHandle, FIONBIO, 0)
    
    'Call the connect function
    If Connect(lngSocketHandle, udtSockaddr, Len(udtSockaddr)) = -1 Then
        Let txtData.Text = txtData.Text & "*** ERROR: Cannot Connect to " & txtHostName.Text & vbCrLf
        Let txtData.SelStart = Len(txtData)
        Call CloseSocket(lngSocketHandle)
        Exit Sub
    Else
        Let txtData.Text = txtData.Text & "*** Connected to " & txtHostName.Text & vbCrLf
        Let txtData.SelStart = Len(txtData)
    End If
    
    'Switch the socket so that it does not block.
    Call ioctlsocket(lngSocketHandle, FIONBIO, 1)
    Let Timer1.Enabled = True
    
End Sub


Private Sub cmdDisconnect_Click()
    'To disconnect, we close the socket
    Call CloseSocket(lngSocketHandle)
    Let lngSocketHandle = 0
    Let txtData.Text = txtData.Text & "*** Disconnected" & vbCrLf
    Let txtData.SelStart = Len(txtData)
End Sub


Private Sub Form_Load()

    If IsIDE Or App.PrevInstance > 0 Then
        'MsgBox "Only run one instance of Visual Basic when using these examples. For some reason blocking sockets block te whole of Visual Basic no matter how many instances are running. Compile these examples and run the exe's to get them working properly."
        'Unload Me
    End If

    'Start Windows sockets version 1
    Call WSAStartup(&H101, WSAInfo)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If lngSocketHandle <> 0 Then Call CloseSocket(lngSocketHandle)
    'Cleanup the Winsock API
    Call WSACleanup
End Sub


Private Sub Timer1_Timer()
  
  Const MAX_BUFFER_LENGTH As Long = 8192

  Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
  Dim lngBytesReceived                    As Long
  Dim strTempBuffer                       As String
    
    'Start looping until we have all the data
    Do
        'Call the recv Winsock API function in order to read data from the buffer
        Let lngBytesReceived = Recv(lngSocketHandle, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
        DoEvents
        
        If lngBytesReceived > 0 Then
        
            'If we have received some data, convert it to the Unicode
            'string that is suitable for the Visual Basic String data type
            Let strTempBuffer = StrConv(arrBuffer, vbUnicode)
    
            'Remove unused bytes
            Let strTempBuffer = Left$(strTempBuffer, lngBytesReceived)
        
            Let txtData.Text = txtData.Text & strTempBuffer & vbCrLf
            Let txtData.SelStart = Len(txtData)
        Else
            'Nothing recieved so exit the loop.
            Exit Do
        End If
    Loop

End Sub


Private Sub txtSend_KeyPress(KeyAscii As Integer)

  Dim arrBuffer()     As Byte
  Dim strData         As String
  Dim BytesSent       As Long
  
    If KeyAscii = vbKeyReturn Then
    
        Let strData = txtSend.Text
        
        'Convert the data string to a byte array
        Let arrBuffer() = StrConv(strData, vbFromUnicode)
        
        'Send the data
        Let BytesSent = Send(lngSocketHandle, arrBuffer(0), Len(strData), 0&)
        
        Let txtData.Text = txtData.Text & "*** " & txtSend.Text & vbCrLf
        Let txtData.SelStart = Len(txtData)
        Call txtSend.SetFocus
        Call SendKeys("{HOME}+{END}")
    End If
End Sub


Private Function IsIDE() As Boolean
  On Error GoTo IDETrue

    Debug.Print 1 / 0
    Exit Function
IDETrue:
    IsIDE = True
End Function
