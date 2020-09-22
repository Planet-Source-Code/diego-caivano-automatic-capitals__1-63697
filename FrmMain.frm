VERSION 5.00
Begin VB.Form FrmAutomaticCapitals 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Automatic Capitals"
   ClientHeight    =   1230
   ClientLeft      =   3645
   ClientTop       =   2070
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4215
   Begin VB.CommandButton CmdSalir 
      Caption         =   "X"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Exit."
      Top             =   680
      Width           =   375
   End
   Begin VB.CommandButton CmdAcerca 
      Caption         =   "?"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "About."
      Top             =   160
      Width           =   375
   End
   Begin VB.Timer TmrHotKeys 
      Interval        =   1
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer TmrCapsLock 
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.CheckBox ChkEstado 
      Caption         =   "Automatic System Activated"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.Shape ShpBorder 
      BorderColor     =   &H0000C000&
      Height          =   975
      Left            =   120
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activate/Deactivate: F11/F12"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   795
      Width           =   2490
   End
End
Attribute VB_Name = "FrmAutomaticCapitals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'HotKeys Declaration
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As KeyCodeConstants) As Long

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VK_CAPITAL = &H14
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'API Declarations
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim X, T, I As Integer

'HotKeys
Private Function KeyDown(ByVal vKey As KeyCodeConstants) As Boolean
    KeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

Private Sub ChkEstado_Click()
    With ChkEstado
        If .Value = 1 Then
            TmrCapsLock.Enabled = True
            .Caption = "Automatic System Activated"
            ShpBorder.BorderColor = &HC000&
        Else
            TmrCapsLock.Enabled = False
            .Caption = "Automatic System Deactivated"
            ShpBorder.BorderColor = vbRed
        End If
    End With
    
End Sub

Private Sub CmdAcerca_Click()
    MsgBox Me.Caption & " ; By Diego Caivano © 2006." & Chr(13) & Chr(13) & _
    "      E-Mail: " & Chr(34) & "dcaivano_ar@hotmail.com" & Chr(34), vbInformation, _
    "About " & Me.Caption
End Sub

Private Sub CmdSalir_Click()
    If CapsLockOn = True Then ToggleCapsLock (False)
    End
End Sub

Private Sub Form_Load()
    'Place Form In The Upper Zone Of The Screen
    Me.Move (Screen.Width - Me.Width) / 2, 60
End Sub

Public Function CapsLockOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyCapital)
    CapsLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function

Public Sub ToggleCapsLock(TurnOn As Boolean)
    'TurnOn (True o False) = CapsLock State
    
    Dim BytKeys(255) As Byte
    Dim bCapsLockOn As Boolean
      
    GetKeyboardState BytKeys(0) 'Get Virtual Key Status
      
    bCapsLockOn = BytKeys(VK_CAPITAL)
    Dim typOS As OSVERSIONINFO
    
    If bCapsLockOn <> TurnOn Then 'If Actual Status <> Required Status
        If typOS.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then 'Windows 95/98
            BytKeys(VK_CAPITAL) = 1
            SetKeyboardState BytKeys(0)
        Else 'Windows NT/2000
            'Simulate Key Press
            keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
            'Simulate Key Release
            keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
    End If
End Sub

Private Sub TmrCapsLock_Timer()
    For I = 1 To 256
        If GetAsyncKeyState(I) = -32767 Then T = I
    Next I
    
    If T = 32 Then '[Space]
        If CapsLockOn = True Or CapsLockOn = False Then ToggleCapsLock (True)
        X = 0 'Reset Digits Counter
    End If
    If T <> 32 And T <> 0 Then
        X = X + 1 'Count Digits
        If X >= 1 And CapsLockOn = True Then
            ToggleCapsLock (False)
        End If
    End If
    
    T = 0
End Sub

Private Sub TmrHotKeys_Timer()
    With ChkEstado
        If KeyDown(vbKeyF11) Then
            .Value = 1
            .Caption = "Automatic System Activated"
            TmrCapsLock.Enabled = True
            ShpBorder.BorderColor = &HC000&
        ElseIf KeyDown(vbKeyF12) Then
            .Value = 0
            .Caption = "Automatic System Deactivated"
            TmrCapsLock.Enabled = False
            ShpBorder.BorderColor = vbRed
        End If
    End With
End Sub
