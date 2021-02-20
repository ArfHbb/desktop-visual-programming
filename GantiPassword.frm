VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GantiPassword.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Batal"
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
      Left            =   3960
      MouseIcon       =   "GantiPassword.frx":C280
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "batal ganti password"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ganti 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ganti"
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
      Left            =   1920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ganti password"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ganti Password"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7335
      Begin VB.TextBox txt_passbaru2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "="
         TabIndex        =   11
         ToolTipText     =   "Konfirmasi password baru"
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txt_passbaru1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "="
         TabIndex        =   10
         ToolTipText     =   "Password baru"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_passlama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "="
         TabIndex        =   9
         ToolTipText     =   "Password lama"
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txt_idpetugas 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         ToolTipText     =   "ID Petugas"
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Password Baru"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Lama"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ID Petugas"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ganti Password"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub bersihwarna()
txt_idpetugas.BackColor = &HFFFFFF
txt_passbaru1.BackColor = &HFFFFFF
txt_passbaru2.BackColor = &HFFFFFF
txt_passlama.BackColor = &HFFFFFF
End Sub

Private Sub cmd_ganti_Click()
bersihwarna
If txt_idpetugas.Text = "" Or txt_passlama.Text = "" Or txt_passbaru1.Text = "" Or txt_passbaru2.Text = "" Then
MsgBox "Maaf tidak boleh ada data kosong !", vbInformation, "Informasi"
bersih
txt_idpetugas.SetFocus
txt_idpetugas.BackColor = &H80FFFF
txt_passbaru1.BackColor = &H80FFFF
txt_passbaru2.BackColor = &H80FFFF
txt_passlama.BackColor = &H80FFFF

ElseIf txt_idpetugas <> iniid Then
MsgBox "ID Petugas tidak sesuai !", vbInformation, "Informasi"
bersih
txt_idpetugas.SetFocus
txt_idpetugas.BackColor = &H80FFFF


ElseIf txt_passlama <> inipass Then
MsgBox "Password lama tidak sesuai !", vbInformation, "Informasi"
bersih
txt_idpetugas.SetFocus
txt_passlama.BackColor = &H80FFFF

ElseIf txt_passbaru1 <> txt_passbaru2 Then
MsgBox "Konfirmasi password baru tidak sesuai !", vbInformation, "Informasi"
bersih
txt_idpetugas.SetFocus
txt_passbaru1.BackColor = &H80FFFF
txt_passbaru2.BackColor = &H80FFFF

Else
Set admino = New ADODB.Recordset
admino.Open ("select * from login where id_petugas like '%" & id & "%' and password_ like '%" & pass & "%'"), koneksi

Dim sqlupdate As String
sqlupdate = "update login set password_='" & txt_passbaru2.Text & "' where id_petugas='" & iniid & "'"
koneksi.Execute sqlupdate
Unload Me
Call bersih
bersihwarna
End If
End Sub
Public Sub bersih()
txt_idpetugas.Text = ""
txt_passlama.Text = ""
txt_passbaru1.Text = ""
txt_passbaru2.Text = ""
End Sub
Private Sub Command1_Click()
Form12.Hide
MDIForm1.Show
End Sub

Private Sub Form_Load()

    
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from login "
Form12.Caption = "Ganti Password"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub



