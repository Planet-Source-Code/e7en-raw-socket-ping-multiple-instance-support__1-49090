VERSION 5.00
Begin VB.UserControl IPtextbox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   1815
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   1785
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.TextBox IP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "0"
         Top             =   20
         Width           =   375
      End
      Begin VB.TextBox IP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "0"
         Top             =   20
         Width           =   375
      End
      Begin VB.TextBox IP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   480
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   20
         Width           =   375
      End
      Begin VB.TextBox IP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   15
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   20
         Width           =   350
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   20
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   20
         Width           =   45
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         Height          =   255
         Left            =   350
         TabIndex        =   5
         Top             =   20
         Width           =   45
      End
   End
End
Attribute VB_Name = "IPtextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Property Get sIP() As String
    sIP = IP(0) & "." & IP(1) & "." & IP(2) & "." & IP(3)
End Property

Property Let sIP(Val As String)
    IP(0) = Split(Val, ".")(0)
    IP(1) = Split(Val, ".")(1)
    IP(2) = Split(Val, ".")(2)
    IP(3) = Split(Val, ".")(3)
End Property

Private Sub IP_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

If Len(IP(Index).Text) = 2 Then
    With IP(Index + 1)
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End If

If KeyAscii = 8 Then Exit Sub
If KeyAscii <= 47 Or KeyAscii >= 58 Then KeyAscii = 0
End Sub

Private Sub IP_LostFocus(Index As Integer)
    If IP(Index).Text = "" Then IP(Index).Text = "0"
    If IP(Index).Text > 255 Then
        
        With IP(Index)
            .SetFocus
            MsgBox sIP & " is a invalid IP", vbCritical + vbApplicationModal, App.Title
            .Text = 255
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
    End If
End Sub
