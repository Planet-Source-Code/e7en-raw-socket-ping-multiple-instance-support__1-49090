VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSocketPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem
Rem|Class Name: cSocketPing.cls                                |Rem
Rem|¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯                                |Rem
Rem|Programmer: Jake Paternoster (§e7eN) <Hate_114@hotmail.com>|Rem
Rem|Date:       8/10/2003                                      |Rem
Rem|                                                           |Rem
Rem| Copyright © 2003 Jake Paternoster <Hate_114@hotmail.com>  |Rem
Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem

Option Explicit

Dim IcmpPacket(0 To 127) As Byte
Dim IcmpReplyPacket(0 To 127)
Dim Checksum(0 To 1) As Byte

Dim lhWnd As Long
Dim lPrevProc As Long
Dim lSocket As Long
Dim lCount As Long
Dim SockAddr As SockAddr_In
Dim sRemoteIP As String
Dim bPinged As Boolean

'==============================================================================
'                                   PROPERTIES
'==============================================================================
Property Let Pinged(Value As String)
    bPinged = Value
End Property

Property Get Pinged() As String
    Pinged = bPinged
End Property

Property Let RemoteIP(Value As String)
    sRemoteIP = Value
End Property

Property Get RemoteIP() As String
    RemoteIP = sRemoteIP
End Property

Property Let Count(Value As Long)
    lCount = Value
End Property

Property Get Count() As Long
    Count = lCount
End Property

Property Let Sock(Value As Long)
    lSocket = Value
End Property

Property Get Sock() As Long
    Sock = lSocket
End Property

Property Let hWnd(Value As Long)
    lhWnd = Value
End Property

Property Get hWnd() As Long
    hWnd = lhWnd
End Property

Property Let PrevProc(Value As Long)
    lPrevProc = Value
End Property

Property Get PrevProc() As Long
    PrevProc = lPrevProc
End Property

'==============================================================================
'                                      SUBS
'==============================================================================

Sub Ping(sIP As String, Optional TimeOut As Long = 3000)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: Creates a Raw ICMP socket and generates a Window handle
'             and window Message. Then assignes it to the socket and
'             hooks the generated window handle. It then creates an
'             ICMP Echo request packet and sends it.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'RFC 792 - Website: http://www.faqs.org/rfcs/rfc792.html
'
'Echo or Echo Reply Message
'
'    0                   1                   2                   3
'    0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1
'   +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'   |     Type      |     Code      |          Checksum             |
'   +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'   |           Identifier          |        Sequence Number        |
'   +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'   |     Data ...
'   +-+-+-+-+-
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-


Dim WSAdata As tagWSAData
Dim Ret As Long
Dim NewhWnd As Long
Dim tCount As Long
Dim x As Integer

If WSAStartup(&H101, WSAdata) = 0 Then

RemoteIP = sIP
Pinged = False

SockAddr.sin_family = AF_INET
SockAddr.sin_port = IPPORT_ECHO
SockAddr.sin_addr = inet_addr(sIP)

If SockAddr.sin_addr = INADDR_NONE Or SockAddr.sin_addr = 0 Then
    MsgBox "Invalid IP!", vbCritical + vbApplicationModal, "Error"
    WSACleanup
    Me.Pinged = True
    Exit Sub
End If

Sock = Socket(AF_INET, SOCK_RAW, IPPROTO_ICMP)
If Sock <> SOCKET_ERROR Then
        If ioctlsocket(Sock, FIONBIO, 1) = 0 Then
        
            hWnd = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
            GenerateMessage
            If WSAAsyncSelect(Sock, hWnd, ICMP_MSG, ByVal FD_READ) = 0 Then
                AddClass Me
                '----------------------------------------
                '           ICMP ECHO PACKET
                '----------------------------------------
                IcmpPacket(0) = 8    'Type: Echo Message
                IcmpPacket(1) = 0    'Code: 0
                IcmpPacket(2) = 0    'Reserved for Checksum
                IcmpPacket(3) = 0    'Reserved for Checksum
                IcmpPacket(4) = &H55 'Identifier
                IcmpPacket(5) = &H55 'Identifier
                IcmpPacket(6) = 0    'Sequence Number
                IcmpPacket(7) = 0    'Sequence Number
                    
                                     'Generate Data to be Echoed
                For x = 8 To 31
                    IcmpPacket(x) = &H41 + x - 8
                Next
                
                CalcCheckSum IcmpPacket, 32, Checksum
        
                IcmpPacket(2) = Checksum(0)    'Reserved for Checksum
                IcmpPacket(3) = Checksum(1)    'Reserved for Checksum
                '----------------------------------------
        
                Ret = sendto(Sock, VarPtr(IcmpPacket(0)), 32, 0, VarPtr(SockAddr), LenB(SockAddr))
                If Ret = -1 Then
                    MsgBox "Winsock Failed SendTo() - " & GetErrorDescription(Err.Number), vbCritical + vbApplicationModal, "Winsock Error"
                End If
                
                Count = GetTickCount
            Else
                MsgBox "Winsock Failed WSAAsyncSelect() - " & GetErrorDescription(Err.Number), vbCritical + vbApplicationModal, "Winsock Error"
            End If
        Else
            MsgBox "Winsock Failed ioctlsocket() - " & GetErrorDescription(Err.Number), vbCritical + vbApplicationModal, "Winsock Error"
        End If
End If
End If

End Sub

Sub ProcessData(Message As Long)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: Checks if there is data to be read from the socket.
'             It then reads if then grabs the reply code and looks it up
'             in reference the message types. Ret return the packet size.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim Ret As Long
Dim Msg As String

Select Case Message
    Case FD_READ

        Ret = recvfrom(Sock, VarPtr(IcmpReplyPacket(0)), 64, 0, VarPtr(SockAddr), LenB(SockAddr))
                    
        If Ret <> -1 Then
            Select Case IcmpReplyPacket(20)
                Case 0
                    Msg = "Echo Reply Received"
                Case 3
                    Msg = "Destination Unreachable"
                Case 4
                    Msg = "Source Quench"
                Case 5
                    Msg = "Redirect"
                Case 8
                    Msg = "Echo"
                Case 11
                    Msg = "Time Exceeded"
                Case 12
                    Msg = "Parameter Problem"
                Case 13
                    Msg = "TimeStamp"
                Case 14
                    Msg = "TimeStamp Reply"
                Case 15
                    Msg = "Information Request"
                Case 16
                    Msg = "Information Reply"
                Case Else
                    Msg = "Unknown Response"
            End Select
            
            Debug.Print Ret
            Pinged = True
            RemoveClass Me
            MsgBox Msg, vbApplicationModal + vbInformation + vbOKOnly, "Success!"
        End If
End Select
End Sub


Sub CalcCheckSum(Pkt() As Byte, PktSize As Long, Checksum() As Byte)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: Generates a Checksum for the ICMP Packet
'
'      Checksum
'      --------
'      The checksum is the 16-bit ones's complement of the one's
'      complement sum of the ICMP message starting with the ICMP Type.
'      For computing the checksum , the checksum field should be zero.
'      If the total length is odd, the received data is padded with one
'      octet of zeros for computing the checksum.  This checksum may be
'      replaced in the future.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim sHex As String
Dim lCheckSum As Long
Dim bByte As Byte
Dim x As Long
Dim y As Long

For x = 0 To PktSize - 1 Step 2
    lCheckSum = lCheckSum + Pkt(x) * 256 + Pkt(x + 1)
Next

lCheckSum = (lCheckSum \ 65535) + (lCheckSum Mod 65536)

sHex = Hex(Not lCheckSum)
sHex = Mid(sHex, Len(sHex) - 3, 4)

For x = 0 To 1
    Checksum(x) = 0
    For y = 1 To 2
        bByte = AscB(Mid(sHex, x * 2 + y, 1)) - &H30
        If bByte > 9 Then
            bByte = bByte - 7
        End If
        Checksum(x) = Checksum(x) * 16 + bByte
    Next
Next
End Sub
