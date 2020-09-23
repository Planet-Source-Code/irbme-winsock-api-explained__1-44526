VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Server"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   315
      Top             =   840
   End
   Begin VB.TextBox txtData 
      Height          =   2220
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   735
      Width           =   5895
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Listening"
      Height          =   540
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   960
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

Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Private Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function CloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long

'Server side Winsock API functions
Private Declare Function Bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function Listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function Accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long
 
'Dat trransfer
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
Private Const SOMAXCONN = &H7FFFFFFF

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
Private lngConnectedSocket As Long


Private Sub cmdStart_Click()

  Dim udtSockaddr As SOCKADDR_IN
  Dim strPort     As String
  Dim intPort     As Integer
  Dim strAddress  As String
  
    'First create a new streat socket that uses the TCP protocol
    Let lngSocketHandle = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    'Error checking
    If lngSocketHandle <= 0 Then
        Call MsgBox("Error creating Socket", vbCritical, "Error!")
        Exit Sub
    End If
    
    'Ask for the port
    Do Until IsNumeric(strPort)
        Let strPort = InputBox("Which port would you like to listen on", "Select a port")
    Loop

    intPort = CInt(strPort)
    
    Let strAddress = InputBox("Which address would you like to bind to", "Select an IP address", "127.0.0.1")
    
    'Fill out the structure
    With udtSockaddr
        Let .sin_addr = inet_addr(strAddress)
        Let .sin_family = AF_INET
        Let .sin_port = htons(intPort)
    End With
    
    Call Bind(lngSocketHandle, udtSockaddr, Len(udtSockaddr))
    
    If MsgBox("Do you wish to listen for connections now? Once you do this the application will appear to freeze until a connection is recieved. Are you sure you wish to continue?", vbYesNo, "Continue?") = vbYes Then
        
        Call Listen(lngSocketHandle, SOMAXCONN)
        Let lngConnectedSocket = Accept(lngSocketHandle, udtSockaddr, Len(udtSockaddr))

        'Switch the socket so that it does not block.
        Call ioctlsocket(lngConnectedSocket, FIONBIO, 1)
        Let Timer1.Enabled = True
    End If
    
    Call CloseSocket(lngSocketHandle)
    
End Sub


Private Sub Form_Load()

    If IsIDE Or App.PrevInstance > 0 Then
        MsgBox "Only run one instance of Visual Basic when using these examples. For some reason blocking sockets block the whole of Visual Basic no matter how many instances are running. Compile these examples and run the exe's to get them working properly."
        Unload Me
    End If
    
    'Start Windows sockets version 1
    Call WSAStartup(&H101, WSAInfo)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If lngSocketHandle <> 0 Then Call CloseSocket(lngSocketHandle)
    If lngConnectedSocket <> 0 Then Call CloseSocket(lngConnectedSocket)
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
        Let lngBytesReceived = Recv(lngConnectedSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
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


Private Function IsIDE() As Boolean
  On Error GoTo IDETrue

    Debug.Print 1 / 0
    Exit Function
IDETrue:
    IsIDE = True
End Function
