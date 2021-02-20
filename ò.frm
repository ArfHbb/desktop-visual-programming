VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Aplikasi Rental VCD dan DVD"
   ClientHeight    =   6285
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12210
   Icon            =   "ò.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "ò.frx":038A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Keterangan"
      Top             =   5910
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   15875
            Text            =   "Rental CD dan DVD HeartDisk"
            TextSave        =   "Rental CD dan DVD HeartDisk"
            Object.ToolTipText     =   "HeartDisk"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "05/01/2017"
            Object.ToolTipText     =   "Tanggal"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12.45"
            Object.ToolTipText     =   "Jam"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Begin VB.Menu mngantipassword 
         Caption         =   "Ganti Password"
      End
      Begin VB.Menu mnLogout 
         Caption         =   "&Logout"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnplay 
         Caption         =   "Play Musik"
      End
      Begin VB.Menu mnpause 
         Caption         =   "Pause Musik"
      End
      Begin VB.Menu mnstop 
         Caption         =   "Stop Musik"
      End
   End
   Begin VB.Menu mnTrans 
      Caption         =   "Transaksi"
      Begin VB.Menu mnPinjam 
         Caption         =   "&Peminjaman"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnKembali 
         Caption         =   "&Pengembalian"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnmaster 
      Caption         =   "Data Master"
      Begin VB.Menu mnAnggota 
         Caption         =   "Pelanggan"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnFilm 
         Caption         =   "Film"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnPetugas 
         Caption         =   "Petugas"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnLapinjam 
         Caption         =   "Laporan Peminjaman"
      End
      Begin VB.Menu mnLapkem 
         Caption         =   "Laporan Pengembalian"
      End
   End
   Begin VB.Menu mnbantuan 
      Caption         =   "Bantuan"
   End
   Begin VB.Menu mnabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const GWL_STYLE = (-16)
Const WS_SYSMENU = &H80000



Private Sub MDIForm_Load()
 Dim l As Long
    l = GetWindowLong(Me.hWnd, GWL_STYLE)
    l = (l And Not WS_SYSMENU)
    l = SetWindowLong(Me.hWnd, GWL_STYLE, l)
    

End Sub

Public Sub UnloadAllForms()
'Dim Form As Form
'For Each Form In Forms
'Unload Form
'Set Form = Nothing
'Next Form
Dim frm As Form

    For Each frm In Forms
        If Not frm.Name = "MDIForm1" Then 'avoid unloading mdi form
            If frm.MDIChild = True Then 'unload only mdi children
                Unload frm
            End If
        End If
    Next frm
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu frmMDIParent.mnMenu, , X, Y
     End If
End Sub


Private Sub mnabout_Click()
UnloadAllForms
MDIForm1.Show
Form13.Show 1
End Sub

Private Sub mnAnggota_Click()
UnloadAllForms
MDIForm1.Hide
Form6.Show
Form6.txt_no_identitas.SetFocus
End Sub

Private Sub mnbantuan_Click()
UnloadAllForms
MDIForm1.Show
Form14.Show 1
End Sub

Private Sub mnExit_Click()
UnloadAllForms
Unload Me
End Sub

Private Sub mnFilm_Click()
UnloadAllForms
MDIForm1.Hide
Form2.Show
Form2.txt_idfilm.SetFocus
End Sub

Private Sub mngantipassword_Click()
MDIForm1.Show
Form12.Show 1
End Sub

Private Sub mnKembali_Click()
UnloadAllForms
MDIForm1.Hide
Form4.Show
Form4.txt_nopinjam.SetFocus
End Sub

Private Sub mnLapinjam_Click()
UnloadAllForms
MDIForm1.Hide
Form5.Show
End Sub

Private Sub mnLapkem_Click()
UnloadAllForms
MDIForm1.Hide
Form8.Show
End Sub

Private Sub mnLogout_Click()
UnloadAllForms
MDIForm1.Hide
Form1.Show
Form1.txt_nama.SetFocus
End Sub

Private Sub mnpause_Click()
frmTip.WindowsMediaPlayer1.Controls.pause
End Sub

Private Sub mnPetugas_Click()
UnloadAllForms
MDIForm1.Hide
Form7.Show
Form7.txt_idptgs.SetFocus
End Sub

Private Sub mnPinjam_Click()
UnloadAllForms
MDIForm1.Hide
Form3.Show
Form3.txt_idpinjam.SetFocus
End Sub

Private Sub mnplay_Click()
frmTip.WindowsMediaPlayer1.Controls.Play
End Sub

Private Sub mnstop_Click()
frmTip.WindowsMediaPlayer1.Controls.Stop
End Sub
