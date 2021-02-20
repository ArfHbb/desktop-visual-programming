VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16440
   Icon            =   "DataPetugas.frx":0000
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   Picture         =   "DataPetugas.frx":000C
   ScaleHeight     =   8430
   ScaleWidth      =   16440
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   960
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Isi Data Petugas"
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
      Height          =   3735
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   5775
      Begin VB.TextBox txt_idptgs 
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
         Left            =   2280
         TabIndex        =   15
         ToolTipText     =   "Masukkan ID Petugas"
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cmb_level 
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
         Height          =   360
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Level petugas"
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txt_passptgs 
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
         Left            =   2280
         TabIndex        =   11
         ToolTipText     =   "Password petugas"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txt_namaptgs 
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
         Left            =   2280
         TabIndex        =   9
         ToolTipText     =   "Nama Petugas"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
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
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data Petugas"
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
      Height          =   4335
      Left            =   9240
      TabIndex        =   1
      Top             =   2520
      Width           =   8415
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
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
         Left            =   5400
         TabIndex        =   18
         ToolTipText     =   "Cetak data petugas"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txt_cari 
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
         Height          =   405
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "Cari data petugas"
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
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
         Left            =   3720
         MaskColor       =   &H00FF0000&
         TabIndex        =   6
         ToolTipText     =   "Hapus data petugas"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H008080FF&
         Caption         =   "Exit"
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
         Left            =   6840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Keluar"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Save"
         Enabled         =   0   'False
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
         Left            =   2160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "simpan perubahan"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         BackColor       =   &H00FF8080&
         Caption         =   "&New"
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
         Left            =   480
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "tambah data baru"
         Top             =   3480
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Pilih data petugas"
         Top             =   720
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cari :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   5280
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Petugas"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pantullllllllllllllllllllllllllllllllllllllllllllllll
Dim pantul As Integer
Private Sub cmb_level_KeyPress(KeyAscii As Integer)
'tidak bisa diidsi angka atau huruf
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub

Private Sub cmd_print_Click()

    CrystalReport1.ReportFileName = App.Path & "\report_data_petugas.rpt"
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1

End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from login order by id_petugas"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

cmb_level.AddItem "admin"
cmb_level.AddItem "owner"
Call bersih


 'pantullllllllllllllllllllllllllllllllllllllllllllllllllllll
 Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

pantul = 100
End Sub

Private Sub cmd_delete_Click()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select*from login where id_petugas='" & txt_idptgs.Text & "'", koneksi
A = rs!id_petugas
If MsgBox("Yakin Ingin Menghapus Data " & A & "?", vbCritical + vbOKCancel, "Hati-hati menghapus data!!") = vbOK Then
Adodc1.Recordset.Delete
DataGrid1.Refresh
End If
Call bersih
End Sub

Private Sub cmd_cancel_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub cmd_save_Click()
If txt_namaptgs.Text = "" Or txt_idptgs.Text = "" Or txt_passptgs.Text = "" Or cmb_level.Text = "" Then
MsgBox "Maaf tidak boleh ada data kosong !", vbInformation, "Informasi"
txt_idptgs.SetFocus
Else

Call BukaDB

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from login"
Adodc1.Refresh
Adodc1.Recordset.Fields("id_petugas") = txt_idptgs.Text
Adodc1.Recordset.Fields("nama_petugas") = txt_namaptgs.Text
Adodc1.Recordset.Fields("password_") = txt_passptgs.Text
Adodc1.Recordset.Fields("level") = cmb_level.Text
Adodc1.Recordset.update
Call bersih

Call bersih
End If
End Sub

Private Sub cmd_new_Click()
If txt_namaptgs.Text = "" Or txt_idptgs.Text = "" Or txt_passptgs.Text = "" Or cmb_level.Text = "" Then
MsgBox "Maaf tidak boleh ada data kosong !", vbInformation, "Informasi"
txt_idptgs.SetFocus
Else

Call BukaDB
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select*from login where id_petugas='" & txt_idptgs.Text & "'", koneksi

If rs.EOF Then 'jika tidak ditemukan, maka isi data
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from login"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("id_petugas") = txt_idptgs.Text
Adodc1.Recordset.Fields("nama_petugas") = txt_namaptgs.Text
Adodc1.Recordset.Fields("password_") = txt_passptgs.Text
Adodc1.Recordset.Fields("level") = cmb_level.Text
Adodc1.Recordset.update
Call bersih

Else 'jika ditemukan, maka ada pesan data sudah ada
A = rs!id_petugas
MsgBox "Data Dengan ID " & A & " Sudah ADA", vbCritical, "SIMPAN"
Call bersih
txt_idptgs.SetFocus
End If ' punya data redundant
End If 'puny data kosong
End Sub
Public Sub bersih()
txt_idptgs.Text = ""
txt_namaptgs.Text = ""
txt_passptgs.Text = ""
End Sub

Private Sub DataGrid1_Click()
txt_idptgs.Text = Adodc1.Recordset!id_petugas
txt_namaptgs.Text = Adodc1.Recordset!nama_petugas
txt_passptgs.Text = Adodc1.Recordset!Password_
cmb_level.Text = Adodc1.Recordset!level
cmd_save.Enabled = True
cmd_delete.Enabled = True
End Sub


Private Sub cari()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select * from login where id_petugas like '%" & txt_cari.Text & "%' or nama_petugas like '%" & txt_cari.Text & "%' or password_ like '%" & txt_cari.Text & "%' or level like '%" & txt_cari.Text & "%'", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub
Private Sub txt_cari_Change()
Call cari
End Sub

'pantullllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllll
Private Sub Timer1_Timer()
Label7.ForeColor = vbBlack
With Label7
 .Left = .Left + pantul
 If .Left < 0 Then pantul = 100
 If .Left > Me.ScaleWidth - .Width Then pantul = -100
 End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub


