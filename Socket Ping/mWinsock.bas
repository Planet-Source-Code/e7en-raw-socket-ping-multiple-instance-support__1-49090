Attribute VB_Name = "mWinsock"
Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem
Rem|Class Name: mWinsock.mod                                   |Rem
Rem|¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯                                   |Rem
Rem|Programmer: Jake Paternoster (§e7eN) <Hate_114@hotmail.com>|Rem
Rem|Date:       8/10/2003                                      |Rem
Rem|                                                           |Rem
Rem| Copyright © 2003 Jake Paternoster <Hate_114@hotmail.com>  |Rem
Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem

Option Explicit
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYSSTATUS_LEN = 256
Public Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Public Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Public Const SOCKET_ERROR = -1
Public Const INADDR_NONE = &HFFFF
Public Const SOCK_RAW = 3
Public Const IPPORT_ECHO = 7
Public Const AF_INET = 2
Public Const IPPROTO_ICMP = 1
Public Const FIONBIO = &H8004667E
Public Const FD_READ = &H1&

Public Const WSABASEERR = 10000
Public Const WSAEINTR = (WSABASEERR + 4)
Public Const WSAEBADF = (WSABASEERR + 9)
Public Const WSAEACCES = (WSABASEERR + 13)
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEMFILE = (WSABASEERR + 24)
Public Const WSAEWOULDBLOCK = (WSABASEERR + 35)
Public Const WSAEINPROGRESS = (WSABASEERR + 36)
Public Const WSAEALREADY = (WSABASEERR + 37)
Public Const WSAENOTSOCK = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ = (WSABASEERR + 39)
Public Const WSAEMSGSIZE = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT = (WSABASEERR + 47)
Public Const WSAEADDRINUSE = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL = (WSABASEERR + 49)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSAENETUNREACH = (WSABASEERR + 51)
Public Const WSAENETRESET = (WSABASEERR + 52)
Public Const WSAECONNABORTED = (WSABASEERR + 53)
Public Const WSAECONNRESET = (WSABASEERR + 54)
Public Const WSAENOBUFS = (WSABASEERR + 55)
Public Const WSAEISCONN = (WSABASEERR + 56)
Public Const WSAENOTCONN = (WSABASEERR + 57)
Public Const WSAESHUTDOWN = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS = (WSABASEERR + 59)
Public Const WSAETIMEDOUT = (WSABASEERR + 60)
Public Const WSAECONNREFUSED = (WSABASEERR + 61)
Public Const WSAELOOP = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH = (WSABASEERR + 65)
Public Const WSAENOTEMPTY = (WSABASEERR + 66)
Public Const WSAEPROCLIM = (WSABASEERR + 67)
Public Const WSAEUSERS = (WSABASEERR + 68)
Public Const WSAEDQUOT = (WSABASEERR + 69)
Public Const WSAESTALE = (WSABASEERR + 70)
Public Const WSAEREMOTE = (WSABASEERR + 71)
Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAEDISCON = (WSABASEERR + 101)
Public Const WSAENOMORE = (WSABASEERR + 102)
Public Const WSAECANCELLED = (WSABASEERR + 103)
Public Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
Public Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)
Public Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
Public Const WSASYSCALLFAILURE = (WSABASEERR + 107)
Public Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
Public Const WSATYPE_NOT_FOUND = (WSABASEERR + 109)
Public Const WSA_E_NO_MORE = (WSABASEERR + 110)
Public Const WSA_E_CANCELLED = (WSABASEERR + 111)
Public Const WSAEREFUSED = (WSABASEERR + 112)
Public Const WSAHOST_NOT_FOUND = (WSABASEERR + 1001)
Public Const WSATRY_AGAIN = (WSABASEERR + 1002)
Public Const WSANO_RECOVERY = (WSABASEERR + 1003)
Public Const WSANO_DATA = (WSABASEERR + 1004)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Type SockAddr_In                 ' IP address struct should be 16 bytes
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Public Type tagWSAData
    wVersion            As Integer
    wHighVersion        As Integer
    szDescription       As String * WSADESCRIPTION_LEN_1
    szSystemStatus      As String * WSASYSSTATUS_LEN_1
    iMaxSockets         As Integer
    iMaxUdpDg           As Integer
    lpVendorInfo        As String * 200
End Type
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function sendto Lib "wsock32.dll" (ByVal Socket As Long, ByVal packet As Long, ByVal packetLen As Long, ByVal flags As Long, ByVal to_addr As Long, ByVal tolen As Long) As Long
Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequested As Integer, lpWSADATA As tagWSAData) As Integer
Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSACleanup Lib "WSOCK32" () As Integer
Public Declare Function closesocket Lib "wsock32.dll" (ByVal Socket As Long) As Long
Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal Socket As Long, ByVal cmd As Long, argp As Long) As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function recvfrom Lib "wsock32.dll" (ByVal Socket As Long, ByVal buffer As Long, ByVal buflen As Long, ByVal flags As Long, ByVal from As Long, fromLen As Long) As Long
Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SockAddr_In, ByRef namelen As Long) As Long
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public ICMP_MSG As Long

'==============================================================================
'                                      SUBS
'==============================================================================

Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: Returns a description of the raised Error
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Dim strDesc As String
    '
    Select Case lngErrorCode
        '
        Case WSAEACCES
            strDesc = "Permission denied."
        Case WSAEADDRINUSE
            strDesc = "Address already in use."
        Case WSAEADDRNOTAVAIL
            strDesc = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
            strDesc = "Address family not supported by protocol family."
        Case WSAEALREADY
            strDesc = "Operation already in progress."
        Case WSAECONNABORTED
            strDesc = "Software caused connection abort."
        Case WSAECONNREFUSED
            strDesc = "Connection refused."
        Case WSAECONNRESET
            strDesc = "Connection reset by peer."
        Case WSAEDESTADDRREQ
            strDesc = "Destination address required."
        Case WSAEFAULT
            strDesc = "Bad address."
        Case WSAEHOSTDOWN
            strDesc = "Host is down."
        Case WSAEHOSTUNREACH
            strDesc = "No route to host."
        Case WSAEINPROGRESS
            strDesc = "Operation now in progress."
        Case WSAEINTR
            strDesc = "Interrupted function call."
        Case WSAEINVAL
            strDesc = "Invalid argument."
        Case WSAEISCONN
            strDesc = "Socket is already connected."
        Case WSAEMFILE
            strDesc = "Too many open files."
        Case WSAEMSGSIZE
            strDesc = "Message too long."
        Case WSAENETDOWN
            strDesc = "Network is down."
        Case WSAENETRESET
            strDesc = "Network dropped connection on reset."
        Case WSAENETUNREACH
            strDesc = "Network is unreachable."
        Case WSAENOBUFS
            strDesc = "No buffer space available."
        Case WSAENOPROTOOPT
            strDesc = "Bad protocol option."
        Case WSAENOTCONN
            strDesc = "Socket is not connected."
        Case WSAENOTSOCK
            strDesc = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
            strDesc = "Operation not supported."
        Case WSAEPFNOSUPPORT
            strDesc = "Protocol family not supported."
        Case WSAEPROCLIM
            strDesc = "Too many processes."
        Case WSAEPROTONOSUPPORT
            strDesc = "Protocol not supported."
        Case WSAEPROTOTYPE
            strDesc = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
            strDesc = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
            strDesc = "Socket type not supported."
        Case WSAETIMEDOUT
            strDesc = "Connection timed out."
        Case WSATYPE_NOT_FOUND
            strDesc = "Class type not found."
        Case WSAEWOULDBLOCK
            strDesc = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
            strDesc = "Host not found."
        Case WSANOTINITIALISED
            strDesc = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
            strDesc = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
            strDesc = "This is a nonrecoverable error."
        Case WSASYSCALLFAILURE
            strDesc = "System call failure."
        Case WSASYSNOTREADY
            strDesc = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
            strDesc = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
            strDesc = "Winsock.dll version out of range."
        Case WSAEDISCON
            strDesc = "Graceful shutdown in progress."
        Case Else
            strDesc = "Unknown error."
    End Select
    '
    GetErrorDescription = strDesc
    '
End Function

Sub GenerateMessage()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Desciption: If a Window Message hasnt been generated then generate one.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If ICMP_MSG = 0 Then
    ICMP_MSG = RegisterWindowMessage("ICMP_Message")
End If
End Sub

