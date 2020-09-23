VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Example 1: Initialising/Terminating"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   435
      Left            =   3045
      TabIndex        =   4
      Top             =   1890
      Width           =   1275
   End
   Begin VB.CommandButton cmdVersion2 
      Caption         =   "Start Version 2.2"
      Height          =   435
      Left            =   1575
      TabIndex        =   3
      Top             =   1890
      Width           =   1380
   End
   Begin VB.TextBox txtValues 
      Height          =   1695
      Left            =   2205
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   105
      Width           =   2115
   End
   Begin VB.TextBox txtParameters 
      Height          =   1695
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   2115
   End
   Begin VB.CommandButton cmdVersion1 
      Caption         =   "Start Version 1.1"
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   1890
      Width           =   1380
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

Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

Private Const SCK_VERSION1 = &H101                  'Windows sockets version 1.1
Private Const SCK_VERSION2 = &H202                  'Windows sockets version 2.2

Private Const OFFSET = 65536

Private Type WSAData
    WVersion        As Integer                      'Version
    WHighVersion    As Integer                      'High Version
    szDescription   As String * WSADESCRIPTION_LEN  'Description
    szSystemStatus  As String * WSASYS_STATUS_LEN   'Status of system
    iMaxSockets     As Integer                      'Maximum number of sockets allowed
    iMaxUdpDg       As Integer                      'Maximum UDP datagrams
    lpVendorInfo    As Long                         'Vendor Info
End Type

Private WSAInfo As WSAData


Private Function IntegerToUnsigned(Value As Integer) As Long

   If Value < 0 Then
       IntegerToUnsigned = Value + OFFSET
   Else
       IntegerToUnsigned = Value
   End If

End Function


Private Sub DisplayInfo(WSAInfo As WSAData)
    
  Dim lngTemp As Long
    
    With WSAInfo
        AddText "Version", Str(.WVersion \ 256 & "." & .WHighVersion Mod 256)
        AddText "High Version", Str(.WHighVersion \ 256 & "." & .WVersion Mod 256)
        lngTemp = IntegerToUnsigned(.iMaxSockets)
        AddText "Maximum sockets", IIf(lngTemp = 0, "Unrestricted", Str(lngTemp))
        lngTemp = IntegerToUnsigned(.iMaxUdpDg)
        AddText "Maximum UDP Datagrams", IIf(lngTemp = 0, "Unrestricted", Str(lngTemp))
        AddText "Vendor information", Str(.lpVendorInfo)
        AddText "Description", Left$(.szDescription, InStr(1, .szDescription, Chr(0)) - 1)
        AddText "System status", Left$(.szSystemStatus, InStr(1, .szSystemStatus, Chr(0)) - 1)
    End With
End Sub


Private Sub AddText(strParameter As String, strValue As String)
    txtParameters.Text = txtParameters.Text & strParameter & vbCrLf
    txtValues.Text = txtValues.Text & strValue & vbCrLf
End Sub


Private Sub cmdQuit_Click()
    Unload Me
End Sub


Private Sub cmdVersion1_Click()
    txtParameters = ""
    txtValues = ""
    Call WSAStartup(SCK_VERSION1, WSAInfo)  'Start the Windows sockets version 1
    Call DisplayInfo(WSAInfo)               'Display the info about it
    Call WSACleanup                         'Terminate the Winsock session
End Sub


Private Sub cmdVersion2_Click()
    txtParameters = ""
    txtValues = ""
    Call WSAStartup(SCK_VERSION2, WSAInfo)  'Start the Windows sockets version 2
    Call DisplayInfo(WSAInfo)               'Display the info about it
    Call WSACleanup                         'Terminate the Winsock session
End Sub

