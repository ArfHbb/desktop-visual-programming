VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16410
   Icon            =   "DataPelanggan.frx":0000
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   Picture         =   "DataPelanggan.frx":000C
   ScaleHeight     =   8430
   ScaleWidth      =   16410
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1320
      Top             =   5760
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
      Caption         =   "Isi Data Pelanggan"
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
      Height          =   4935
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   6015
      Begin VB.ComboBox cbo_status 
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
         Left            =   2520
         TabIndex        =   18
         ToolTipText     =   "Masukkan sattus keanggotaan"
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txt_no_identitas 
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
         Left            =   2520
         TabIndex        =   16
         ToolTipText     =   "KTP/KTM/SIM/Nomor Anggota"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txt_telp 
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
         Left            =   2520
         TabIndex        =   14
         ToolTipText     =   "Nomor telepon"
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox txt_alamat 
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
         Left            =   2520
         TabIndex        =   12
         ToolTipText     =   "Alamat pelanggan"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txt_nmplgn 
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
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "Nama Lengkap"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lbl_idpelanggan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label4"
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
         Left            =   2520
         TabIndex        =   20
         ToolTipText     =   "ID masing-masing Pelanggan"
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Identitas"
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
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Telp"
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
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label Label5 
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
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ID Pelanggan"
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
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data Pelanggan"
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
      Height          =   5775
      Left            =   8640
      TabIndex        =   1
      Top             =   2160
      Width           =   7335
      Begin VB.CommandButton cmd_cetak 
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
         Left            =   4560
         TabIndex        =   22
         ToolTipText     =   "Cetak data pelanggan"
         Top             =   5040
         Width           =   1095
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
         TabIndex        =   19
         ToolTipText     =   "Cari data anggota"
         Top             =   240
         Width           =   3735
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Keluar"
         Top             =   5040
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
         Left            =   240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Tambah data baru"
         Top             =   5040
         Width           =   1215
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
         Left            =   3120
         MaskColor       =   &H00FF0000&
         TabIndex        =   4
         ToolTipText     =   "Hapus data pelanggan"
         Top             =   5040
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
         Left            =   1680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Simpan perubahan"
         Top             =   5040
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4095
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Pilih data anggita"
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         Enabled         =   -1  'True
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
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   7440
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
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
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pelanggan"
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
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   840
      Width           =   8055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pantullllllllllllllllllllllllllllllllllllllllllllllll
Dim pantul As Integer
Private Sub cbo_status_KeyPress(KeyAscii As Integer)
'tidak bisa diidsi angka atau huruf
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub

Private Sub cmd_cetak_Click()
CrystalReport1.ReportFileName = App.Path & "\report_data_pelanggan.rpt"
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from anggota order by id_Anggota"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
cbo_status.AddItem "Member"
cbo_status.AddItem "Non-member"
cbo_status.AddItem "Non-aktif"
Call KodeOtomatis
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
rs.Open "select*from anggota where no_identitas='" & txt_no_identitas.Text & "'", koneksi
A = rs!no_identitas
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
If txt_no_identitas.Text = "" Or txt_nmplgn.Text = "" Or txt_alamat = "" Or txt_telp = "" Or cbo_status.Text = "" Then
MsgBox "Maaf tidak boleh ada data kosong !", vbInformation, "Informasi"
txt_no_identitas.SetFocus
Else

Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from anggota"
Adodc1.Refresh

Adodc1.Recordset.Fields("no_identitas") = txt_no_identitas.Text
Adodc1.Recordset.Fields("Nama") = txt_nmplgn.Text
Adodc1.Recordset.Fields("Alamat") = txt_alamat.Text
Adodc1.Recordset.Fields("Telephone") = txt_telp.Text
Adodc1.Recordset.Fields("status") = cbo_status.Text
Adodc1.Recordset.update
Call bersih
End If
End Sub

Private Sub cmd_new_Click()
If txt_no_identitas.Text = "" Or txt_nmplgn.Text = "" Or txt_alamat = "" Or txt_telp = "" Or cbo_status.Text = "" Then
MsgBox "Maaf tidak boleh ada data kosong !", vbInformation, "Informasi"
txt_no_identitas.SetFocus
Else

Call KodeOtomatis
Call BukaDB
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select*from anggota where no_identitas='" & txt_no_identitas.Text & "'", koneksi

If rs.EOF Then 'jika tidak ditemukan, maka isi data
Call BukaDB

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from anggota"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("id_anggota") = lbl_idpelanggan.Caption
Adodc1.Recordset.Fields("no_identitas") = txt_no_identitas.Text
Adodc1.Recordset.Fields("Nama") = txt_nmplgn.Text
Adodc1.Recordset.Fields("Alamat") = txt_alamat.Text
Adodc1.Recordset.Fields("Telephone") = txt_telp.Text
Adodc1.Recordset.Fields("status") = cbo_status.Text
Adodc1.Recordset.update

Else 'jika ditemukan, maka ada pesan data sudah ada
A = rs!no_identitas
MsgBox "Pelanggan dengan nomor identitas " & A & " Sudah ADA", vbCritical, "SIMPAN"
txt_no_identitas.SetFocus

End If
Call bersih
End If
End Sub

Private Sub DataGrid1_Click()
txt_no_identitas.Text = Adodc1.Recordset!no_identitas
txt_nmplgn.Text = Adodc1.Recordset!nama
txt_alamat.Text = Adodc1.Recordset!alamat
txt_telp.Text = Adodc1.Recordset!telephone
cbo_status.Text = Adodc1.Recordset!status
cmd_save.Enabled = True
cmd_delete.Enabled = True
End Sub
Public Sub bersih()
txt_no_identitas.Text = ""
txt_nmplgn.Text = ""
txt_alamat.Text = ""
txt_telp.Text = ""
End Sub

Sub KodeOtomatis()
Call BukaDB
admin.Open ("select * from anggota Where id_anggota In(Select Max(id_anggota)From anggota)Order By id_anggota Desc"), koneksi
admin.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With admin
        If .EOF Then
            Urutan = "PLG" + "001"
            lbl_idpelanggan = Urutan
        Else
            Hitung = Right(!id_anggota, 3) + 1
            Urutan = "PLG" + Right("000" & Hitung, 3)
        End If
        lbl_idpelanggan = Urutan
    End With
End Sub
Private Sub cari()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select * from anggota where id_Anggota like '%" & txt_cari.Text & "%' or no_identitas like '%" & txt_cari.Text & "%' or Nama like '%" & txt_cari.Text & "%' or Alamat like '%" & txt_cari.Text & "%' or Telephone like '%" & txt_cari.Text & "%' or Status like '%" & txt_cari.Text & "%'", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

Private Sub txt_cari_Change()
Call cari
End Sub

'pantullllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllll
Private Sub Timer1_Timer()
Label9.ForeColor = vbBlack
With Label9
 .Left = .Left + pantul
 If .Left < 0 Then pantul = 100
 If .Left > Me.ScaleWidth - .Width Then pantul = -100
 End With
End Sub

Private Sub txt_nmplgn_Change()
If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0

End Sub

Private Sub txt_telp_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub


