VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Socket Ping"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "IP Address"
      Height          =   615
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdPing 
         Caption         =   "Ping"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin Project1.IPtextbox IPtxt 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Date:         8/10/2003
'Progammer:    Jake Paternoster (§e7eN)
'Email:        Hate_114@hotmail.com
'Program Name: Socket Ping
                
'Description:  This is just a Simple example of pinging using RAW sockets.
'              The idea of this code it to be used in multi-instancing of
'              ICMP requests.
'
'              Please Vote and Comment.

Option Explicit

Private Declare Function GetInputState Lib "user32.dll" () As Long
Private Const TimeOut = 3000

Dim cSP As New cSocketPing

Private Sub cmdPing_Click()
    cSP.Ping IPtxt.sIP
    
    Do Until GetTickCount() - cSP.Count > TimeOut Or cSP.Pinged = True
        If GetInputState = 0 Then DoEvents
    Loop
    
    If cSP.Pinged = False Then
        RemoveClass cSP
        MsgBox "Request Timed Out!", vbApplicationModal + vbCritical + vbOKOnly, "Timed-out"
    End If

End Sub

