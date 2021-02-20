VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informasi Program"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7320
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Informasi"
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   6735
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "About.frx":C280
         ToolTipText     =   "Infromasi tentang kami"
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label LBL_INSTAGRAM 
         BackStyle       =   0  'Transparent
         Caption         =   "INSTAGRAM"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         ToolTipText     =   "Lihat instagram kami"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lbl_twitter 
         BackStyle       =   0  'Transparent
         Caption         =   "TWITTER"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         ToolTipText     =   "Lihat twitter kami"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lbl_fb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FACEBOOK"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         ToolTipText     =   "Lihat Facebook kami"
         Top             =   4800
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmd_tutup 
      BackColor       =   &H008080FF&
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "tutup"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal h As Long, ByVal hb As Long, ByVal X As Long, _
ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal F As Long) As Long

Const SWP_NOMOVE = 2

Const SWP_NOSIZE = 1

Const flags = SWP_NOMOVE Or SWP_NOSIZE

Const HWND_TOPMOST = -1

Const HWND_NOTOPMOST = -2



Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const GWL_STYLE = (-16)
Const WS_SYSMENU = &H80000

 

Private Sub cmd_tutup_Click()
Form13.Hide
MDIForm1.Show
End Sub

Private Sub Form_Load()
 Dim l As Long
    l = GetWindowLong(Me.hWnd, GWL_STYLE)
    l = (l And Not WS_SYSMENU)
    l = SetWindowLong(Me.hWnd, GWL_STYLE, l)
    
      
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub

Private Sub lbl_fb_Click()
ShellExecute Me.hWnd, "Open", "http://facebook.com/heartdisk", vbNullString, vbNullString, 3
End Sub

Private Sub LBL_INSTAGRAM_Click()
ShellExecute Me.hWnd, "Open", "http://INSTAGRAM.com/heartdisk", vbNullString, vbNullString, 3
End Sub

Private Sub lbl_twitter_Click()
ShellExecute Me.hWnd, "Open", "http://twitter.com/heartdisk", vbNullString, vbNullString, 3
End Sub

