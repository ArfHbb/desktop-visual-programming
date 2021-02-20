VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form2 
   Caption         =   "Data Film"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17010
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2.frx":000C
   ScaleHeight     =   8430
   ScaleWidth      =   17010
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7560
      Top             =   2040
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
      Caption         =   "Data Film"
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
      Height          =   5295
      Left            =   9120
      TabIndex        =   8
      Top             =   2040
      Width           =   7695
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
         Left            =   4800
         TabIndex        =   18
         ToolTipText     =   "Cetak data film"
         Top             =   4440
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
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         ToolTipText     =   "Cari data film"
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H008080FF&
         Caption         =   "&Cancel"
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
         Left            =   6120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Keluar"
         Top             =   4440
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
         Left            =   3240
         MaskColor       =   &H00FF0000&
         Picture         =   "Form2.frx":F907
         TabIndex        =   12
         ToolTipText     =   "Hapus data film"
         Top             =   4440
         Width           =   1335
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
         Left            =   1800
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Simpan perubahan"
         Top             =   4440
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
         Left            =   360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Tambah data baru"
         Top             =   4440
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3255
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "Pilih film"
         Top             =   840
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         Caption         =   "Film"
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
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Isi Data Film"
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
      Height          =   4695
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   6135
      Begin VB.TextBox txt_stok 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         ToolTipText     =   "Stok Film"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.ComboBox cbo_iddisk 
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
         TabIndex        =   17
         ToolTipText     =   "1 : VCD 2: DVD"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox cbo_harga 
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
         TabIndex        =   16
         ToolTipText     =   "VCD : RP 3000 DVD : RP 5000"
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txt_judul 
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
         Height          =   855
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Masukkan judul film"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txt_idfilm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   3
         ToolTipText     =   "Masukkan ID Film"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stok"
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
         Left            =   480
         TabIndex        =   19
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Disk"
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
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga (Rp)"
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
         Left            =   480
         TabIndex        =   6
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Judul"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Film"
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
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   5520
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
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
   Begin VB.Image Image1 
      Height          =   735
      Left            =   5520
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Film"
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
      Left            =   6240
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pantullllllllllllllllllllllllllllllllllllllllllllllll
Dim pantul As Integer
Private Sub addcombo()
Call BukaDB
    Set RsDataCombo = New ADODB.Recordset
    RsDataCombo.Open "SELECT * FROM disk", koneksi, adOpenDynamic, adLockOptimistic
    Do Until RsDataCombo.EOF
        cbo_iddisk.AddItem RsDataCombo!id_disk
        RsDataCombo.MoveNext
    Loop
    RsDataCombo.Close
End Sub

Private Sub comboharga()
Call BukaDB
    Set RsDataCombo = New ADODB.Recordset
    RsDataCombo.Open "SELECT * FROM disk", koneksi, adOpenDynamic, adLockOptimistic
    Do Until RsDataCombo.EOF
        cbo_harga.AddItem RsDataCombo!harga_sewa
        RsDataCombo.MoveNext
    Loop
    RsDataCombo.Close
End Sub

Private Sub cbo_harga_KeyPress(KeyAscii As Integer)
'tidak bisa diidsi angka atau huruf
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub


Private Sub cbo_iddisk_KeyPress(KeyAscii As Integer)
'tidak bisa diidsi angka atau huruf
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub cmd_print_Click()

    CrystalReport1.ReportFileName = App.Path & "\report_data_film.rpt"
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from film order by id_film"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
With DataGrid1
.Columns(0).Caption = "ID Film"
.Columns(1).Caption = "ID Disk"
.Columns(2).Caption = "Judul"
.Columns(3).Caption = "Harga"

.Columns(1).Alignment = dbgCenter
.Columns(3).Alignment = dbgCenter
End With
DataGrid1.Refresh
Call addcombo
Call comboharga
 Call bersih

 
 
 'pantullllllllllllllllllllllllllllllllllllllllllllllllllllll
 Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

pantul = 100

txt_idfilm.MaxLength = 6
txt_stok.MaxLength = 1
End Sub
Private Sub cmd_delete_Click()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select*from film where id_film='" & txt_idfilm.Text & "'", koneksi
A = rs!id_film
If MsgBox("Yakin Ingin Menghapus Data " & A & "?", vbCritical + vbOKCancel, "Hati-hati menghapus data!!") = vbOK Then
Adodc1.Recordset.Delete
DataGrid1.Refresh
End If
Call bersih
End Sub

Private Sub cmd_cancel_Click()
Form2.Hide
MDIForm1.Show
End Sub

Private Sub cmd_save_Click()
If txt_idfilm.Text = "" Or cbo_iddisk.Text = "" Or txt_judul.Text = "" Or cbo_harga = "" Then
MsgBox "Maaf tidak boleh ada data kosong !", vbInformation, "Informasi"
Else

Adodc1.Recordset.Fields("id_film") = txt_idfilm.Text
Adodc1.Recordset.Fields("Judul") = txt_judul.Text
Adodc1.Recordset.Fields("Harga") = cbo_harga.Text
Adodc1.Recordset.Fields("id_disk") = cbo_iddisk.Text
Adodc1.Recordset.Fields("jumlah_stok") = txt_stok.Text
Adodc1.Recordset.update
Call bersih
End If
End Sub

Private Sub cmd_new_Click()
If txt_idfilm.Text = "" Or cbo_iddisk.Text = "" Or txt_judul.Text = "" Or cbo_harga = "" Then
MsgBox "Maaf tidak boleh ada data kosong !", vbInformation, "Informasi"
Else 'jika data tidak kosong

Call BukaDB
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
    
rs.CursorLocation = adUseClient
rs.Open "select*from film where id_film='" & txt_idfilm.Text & "'", koneksi

If rs.EOF Then 'jika tdk ditemukan, maka isi data
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from film"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("id_film") = txt_idfilm.Text
Adodc1.Recordset.Fields("Judul") = txt_judul.Text
Adodc1.Recordset.Fields("Harga") = cbo_harga.Text
Adodc1.Recordset.Fields("id_disk") = cbo_iddisk.Text
Adodc1.Recordset.Fields("jumlah_stok") = txt_stok.Text
Adodc1.Recordset.update

Else 'jika  ditemukan, maka isi data gak bisa

A = rs!id_film
MsgBox "Data Dengan ID Film " & A & "Sudah ADA", vbCritical, "SIMPAN"

Call bersih
End If 'punya data ganda
End If 'punya data kosong
End Sub
Private Sub DataGrid1_Click()
txt_idfilm.Text = Adodc1.Recordset!id_film
txt_judul.Text = Adodc1.Recordset!judul
cbo_harga.Text = Adodc1.Recordset!Harga
cbo_iddisk.Text = Adodc1.Recordset!id_disk
txt_stok.Text = Adodc1.Recordset!Jumlah_stok
cmd_delete.Enabled = True
cmd_save.Enabled = True
End Sub
Public Sub bersih()
txt_idfilm.Text = ""
txt_judul.Text = ""
End Sub
Private Sub cari()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select * from film where id_film like '%" & txt_cari.Text & "%' or id_disk like '%" & txt_cari.Text & "%' or Judul like '%" & txt_cari.Text & "%' or Harga like '%" & txt_cari.Text & "%'", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
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

Private Sub txt_cari_Change()
Call cari
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub


