VERSION 5.00
Begin VB.Form MemBar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2160
      Top             =   1440
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "MemBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
'-----------------------------------------------------------------------
Private Type MEMORYSTATUS
   dwLength As Long
   dwMemoryLoad As Long
  dwTotalPhys As Long
  dwAvailPhys As Long
  dwTotalPageFile As Long
  dwAvailPageFile As Long
  dwTotalVirtual As Long
  dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'-----------------------------------------------------------------------
Dim memInfo As MEMORYSTATUS

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub Form_Load()
    Me.Width = 2300
    Me.Height = 970
    Me.Caption = "Physical/Page/Virtual"
End Sub

Private Sub Form_Resize()
    Dim w, h As Integer
    w = Me.Width
    h = Me.Height
    Dim i As Integer
    For i = 0 To (3 * 3 - 1)
        lbl(i).Top = 0
        lbl(i).Left = 0
        If i > 2 Then lbl(i).Top = lbl(i - 3).Top + lbl(i - 3).Height - 8
        lbl(i).Width = w - 120
        lbl(i).Height = (h - 360) / 3
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MemBar = Nothing
    End
End Sub

Private Sub lbl_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    MakeTopMost Me.hwnd
    GlobalMemoryStatus memInfo
    
    high = byteToMB(memInfo.dwTotalPhys)
    current = byteToMB(memInfo.dwAvailPhys)
    lbl(1).Width = ((high - current) / high) * lbl(0).Width
    lbl(2).Caption = "Free: " & current & "MB (" & 100 - Round((high - current) / high * 100, 0) & "%)"
   
    high = byteToMB(memInfo.dwTotalPageFile)
    current = byteToMB(memInfo.dwAvailPageFile)
    lbl(4).Width = ((high - current) / high) * lbl(0).Width
    lbl(5).Caption = "Free: " & current & "MB (" & 100 - Round((high - current) / high * 100, 0) & "%)"
    
    high = byteToMB(memInfo.dwTotalVirtual)
    current = byteToMB(memInfo.dwAvailVirtual)
    lbl(7).Width = ((high - current) / high) * lbl(0).Width
    lbl(8).Caption = "Free: " & current & "MB (" & 100 - Round((high - current) / high * 100, 0) & "%)"
    
    
End Sub

Private Function byteToMB(ByVal b As Long)
    byteToMB = Round(((b / 1024) / 1024), 2)
End Function

