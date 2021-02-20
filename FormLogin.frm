VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   4935
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11280
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormLogin.frx":014A
   ScaleHeight     =   4935
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   -1080
      Picture         =   "FormLogin.frx":4A9D
      ScaleHeight     =   6435
      ScaleWidth      =   13155
      TabIndex        =   0
      Top             =   -360
      Width           =   13215
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Masuk"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   4440
         TabIndex        =   2
         Top             =   960
         Width           =   5415
         Begin VB.CommandButton cmd_batal 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "Batal"
            DisabledPicture =   "FormLogin.frx":EFA9
            DownPicture     =   "FormLogin.frx":138FC
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmd_login 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            Caption         =   "Login"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox txt_pass 
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
            ForeColor       =   &H00FF0000&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   2280
            PasswordChar    =   "="
            TabIndex        =   8
            Top             =   2760
            Width           =   2055
         End
         Begin VB.TextBox txt_nama 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   2280
            TabIndex        =   7
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
            Height          =   15
            Left            =   2760
            TabIndex        =   14
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lbl_pass 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1200
            TabIndex        =   9
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label lbl_nama 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID :"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1200
            TabIndex        =   6
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Untuk melakukan transaksi, silahkan Login terlebih dahulu!"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   1800
            Width           =   4815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Anda sedang berada di aplikasi rental VCD dan DVD. "
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   600
            TabIndex        =   4
            Top             =   1440
            Width           =   4335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Selamat Datang di HEARTDISK"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Left            =   360
            TabIndex        =   3
            Top             =   120
            Width           =   5175
         End
      End
      Begin VB.Label pass 
         Caption         =   "ganti password.frm"
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label id 
         Caption         =   "jgn dihapus"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Label4"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Label1"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   615
      End
      Begin VB.Label status 
         Caption         =   "jgn dihapus"
         Height          =   375
         Left            =   11400
         TabIndex        =   1
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Menu mnoperasi 
      Caption         =   "operasi"
      Visible         =   0   'False
      Begin VB.Menu mnlogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnbatal 
         Caption         =   "Batal"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const GWL_STYLE = (-16)
Const WS_SYSMENU = &H80000


Public Sub bersihwarna()
txt_nama.BackColor = &H8000000F
txt_pass.BackColor = &H8000000F
End Sub
Public Sub bersih()
            txt_nama = ""
            txt_pass = ""
End Sub

Private Sub cmd_batal_Click()
Form1.Hide
Unload Me
End Sub
Private Sub cmd_login_Click()

Dim KodeAdmin As String
Dim NamaAdmin As String
Call BukaDB
Dim sql As String
Dim cari As New ADODB.Recordset
sql = "select * from login where id_petugas ='" & LCase(txt_nama.Text) & "' and password_ =  '" & LCase(txt_pass.Text) & "'"
Set cari = koneksi.Execute(sql)

If cari.EOF Then
   If txt_nama.Text = "" And txt_pass.Text = "" Then
   bersihwarna
MsgBox "Maaf nama dan password anda kosong. Harap diisi !", vbInformation, "Informasi"
bersih
txt_nama.SetFocus
txt_nama.BackColor = &H80FFFF
txt_pass.BackColor = &H80FFFF
ElseIf txt_nama.Text = "" Then
bersihwarna
MsgBox "Maaf nama tidak boleh kosong !", vbInformation, "Informasi"
txt_nama.SetFocus
txt_nama.BackColor = &H80FFFF
ElseIf txt_pass.Text = "" Then
bersihwarna
MsgBox "Maaf password tidak boleh kosong !", vbInformation, "Informasi"
txt_pass.SetFocus
txt_pass.BackColor = &H80FFFF
Else

bersihwarna
sql2 = "select id_petugas from login where id_petugas ='" & LCase(txt_nama.Text) & "'"
Set cari2 = koneksi.Execute(sql2)
sql3 = "select password_ from login where password_ ='" & LCase(txt_pass.Text) & "'"
Set cari3 = koneksi.Execute(sql3)

If cari2.EOF And cari3.EOF Then
bersihwarna
MsgBox "Maaf User dan password anda salah !", vbInformation, "Informasi"

   bersih
   txt_nama.SetFocus
   txt_nama.BackColor = &H80FFFF
   txt_pass.BackColor = &H80FFFF


      
ElseIf cari2.EOF Then
bersihwarna
MsgBox "Maaf User anda salah !", vbInformation, "Informasi"
   bersih
   txt_nama.SetFocus
txt_nama.BackColor = &H80FFFF

ElseIf cari3.EOF Then
MsgBox "Maaf password anda salah !", vbInformation, "Informasi"
   bersih
   txt_pass.SetFocus
txt_pass.BackColor = &H80FFFF
   
    End If
    
    End If
    
Else
bersihwarna

Dim level As String
status.Caption = cari!level
 MsgBox "Anda berhasil login !", vbInformation, "Informasi"
 If status.Caption = "admin" Then
    MDIForm1.mnLaporan.Visible = False
    MDIForm1.mnPetugas.Visible = False
ElseIf status.Caption = "owner" Then
    MDIForm1.mnLaporan.Visible = True
    MDIForm1.mnPetugas.Visible = True
End If
iniid = txt_nama.Text
inipass = txt_pass.Text
 frmTip.Show
 Form1.Hide
 
End If
End Sub

Private Sub Form1_Load()
 Dim l As Long
    l = GetWindowLong(Me.hwnd, GWL_STYLE)
    l = (l And Not WS_SYSMENU)
    l = SetWindowLong(Me.hwnd, GWL_STYLE, l)

Label4.Caption = Format(Now, "long date")
Label8 = Date

Call bersih


 End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnoperasi
End Sub



Private Sub txt_nama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txt_pass.SetFocus
End If
End Sub

Private Sub txt_pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmd_login_Click
End If
End Sub



