Attribute VB_Name = "mSubclass"
Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem
Rem|Class Name: mSubclass.mod                                  |Rem
Rem|¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯                                  |Rem
Rem|Programmer: Jake Paternoster (§e7eN) <Hate_114@hotmail.com>|Rem
Rem|Date:       8/10/2003                                      |Rem
Rem|                                                           |Rem
Rem| Copyright © 2003 Jake Paternoster <Hate_114@hotmail.com>  |Rem
Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem-Rem

Option Explicit
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Const GWL_WNDPROC As Long = -4&
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Dim SocketPings As Collection
Dim lPrevProc As Long

'==============================================================================
'                                      SUBS
'==============================================================================

Sub AddClass(SocketPing As cSocketPing)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: Adds a class to to a collection so we can send stuff to later.
'             Also Hooks the hWnd generated when we sent the Ping.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    If SocketPings Is Nothing Then
        Set SocketPings = New Collection
    End If
    
    If SocketPing.hWnd <> 0 Then
        lPrevProc = SetWindowLong(SocketPing.hWnd, GWL_WNDPROC, AddressOf WindowProc)
        SocketPing.PrevProc = lPrevProc
        SocketPings.Add SocketPing
    End If
End Sub

Sub RemoveClass(SocketPing As cSocketPing)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: Removes the Class from the collection and UnHooks the hWnd.
'             Also Closes the socket and destroys the Generated window Handle
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim x As Integer

If Not SocketPings Is Nothing Then
If SocketPings.Count <> 0 Then
    For x = 1 To SocketPings.Count
        If SocketPings(x).hWnd = SocketPing.hWnd Then
            SetWindowLong SocketPings(x).hWnd, GWL_WNDPROC, SocketPings(x).PrevProc
            closesocket SocketPings(x).Sock
            DestroyWindow SocketPings(x).hWnd
            WSACleanup
            SocketPings.Remove x
            Exit For
        End If
    Next
End If
End If
End Sub

Function ReturnClass(Optional hWnd As Long = -1, Optional Socket As Long = -1) As cSocketPing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: Returns matching class matching the Socket and hWnd
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim x As Integer
Dim TmpSocketPing As cSocketPing
    For x = 1 To SocketPings.Count
        Set TmpSocketPing = SocketPings(x)
        
        If hWnd <> -1 Then
            If TmpSocketPing.hWnd = hWnd Then
                Set ReturnClass = TmpSocketPing
                Exit For
            End If
        End If
        
        If Socket <> -1 Then
            If TmpSocketPing.Sock = Socket Then
                Set ReturnClass = TmpSocketPing
                Exit For
            End If
        End If
    Next
End Function

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Description: This is where all the hooked messages are sent. If the message
'             received matches our generated Icmp Message then its from one
'             of our sockets so get which class its from and send it the message
'
'             hWnd = Window Handle
'             uMsg = Window Message
'             wParam = Socket
'             lParam = Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim TmpSocketPing As cSocketPing
DoEvents

If uMsg = ICMP_MSG Then
    Set TmpSocketPing = ReturnClass(hWnd, wParam)
    
        If Not TmpSocketPing Is Nothing Then
            TmpSocketPing.ProcessData lParam
        Else
            WindowProc = CallWindowProc(lPrevProc, hWnd, uMsg, wParam, lParam)
        End If
Else
    'Set TmpSocketPing = ReturnClass(hWnd)
    WindowProc = CallWindowProc(lPrevProc, hWnd, uMsg, wParam, lParam)
End If

End Function

