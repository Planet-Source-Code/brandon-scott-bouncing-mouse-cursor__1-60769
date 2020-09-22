VERSION 5.00
Begin VB.Form frmBouncingMouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bouncing Mouse - By Xeon"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4440
      Top             =   120
   End
   Begin VB.Label lblInfo 
      Caption         =   "This is an amazing feat of physics, enough time, and completely nothing to do!"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmBouncingMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim XState As Integer
Dim YState As Integer

Private Sub Timer1_Timer()
    
    Const Speed As Integer = 10

    Dim ptCurrentPosition As POINTAPI
    Call GetCursorPos(ptCurrentPosition)
    
    'lblMouseInfo.Caption = "X=" & ptCurrentPosition.X & " Y=" & ptCurrentPosition.Y
    
    If XState = 0 Then
        If (ptCurrentPosition.X + 1) >= (Screen.Width \ Screen.TwipsPerPixelX) Then
            ptCurrentPosition.X = ptCurrentPosition.X - Speed
            Beep 3000, 100
            XState = 1
        Else
            ptCurrentPosition.X = ptCurrentPosition.X + Speed
        End If
    Else
        If ptCurrentPosition.X <= 0 Then
            ptCurrentPosition.X = ptCurrentPosition.X + Speed
            Beep 3000, 100
            XState = 0
        Else
            ptCurrentPosition.X = ptCurrentPosition.X - Speed
        End If
    End If
    
    If YState = 0 Then
        If (ptCurrentPosition.Y + 1) >= (Screen.Height \ Screen.TwipsPerPixelY) Then
            ptCurrentPosition.Y = ptCurrentPosition.Y - Speed
            Beep 3000, 100
            YState = 1
        Else
            ptCurrentPosition.Y = ptCurrentPosition.Y + Speed
        End If
    Else
        If ptCurrentPosition.Y <= 0 Then
            ptCurrentPosition.Y = ptCurrentPosition.Y + Speed
            Beep 3000, 100
            YState = 0
        Else
            ptCurrentPosition.Y = ptCurrentPosition.Y - Speed
        End If
    End If
    
    Call SetCursorPos(ptCurrentPosition.X, ptCurrentPosition.Y)
End Sub
