VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form5 
   Caption         =   "Laporan Peminjaman"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14550
   Icon            =   "Laporan.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   Picture         =   "Laporan.frx":000C
   ScaleHeight     =   7875
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2760
      Top             =   3000
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12360
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Data Laporan Peminjaman"
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
      Left            =   5640
      TabIndex        =   1
      Top             =   1920
      Width           =   8655
      Begin VB.OptionButton opt_cetakstruk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cetak Struk"
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
         Height          =   255
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   11
         ToolTipText     =   "Cetak semua data"
         Top             =   4320
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   4920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51838977
         CurrentDate     =   42730
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   4920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51838977
         CurrentDate     =   42730
      End
      Begin VB.OptionButton opt_cetakwaktu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cetak dari tanggal"
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
         Height          =   255
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   8
         ToolTipText     =   "cetak data berdasarkan tanggal"
         Top             =   4920
         Width           =   2415
      End
      Begin VB.OptionButton opt_semua 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cetak Semua"
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
         Height          =   255
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   7
         ToolTipText     =   "Cetak semua data"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton cmd_cancel 
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
         Left            =   6480
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Batal (Kembali)"
         Top             =   3480
         Width           =   1575
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
         TabIndex        =   4
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmd_cetak 
         BackColor       =   &H00FF8080&
         Caption         =   "Cetak"
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
         Left            =   4680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cetak data peminjaman"
         Top             =   3480
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2535
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4471
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
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Peminjaman"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   960
      Width           =   8055
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pantullllllllllllllllllllllllllllllllllllllllllllllll
Dim pantul As Integer
Private Sub cmd_cancel_Click()
Form5.Hide
Unload Me
End Sub

Private Sub cmd_cetak_Click()
'salah pilih report, yang peminjaman belum dibuat
If opt_semua.Value = True Then
CrystalReport1.ReportFileName = App.Path & "\laporan_peminjaman.rpt"
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
ElseIf opt_cetakwaktu.Value = True Then
CrystalReport1.SelectionFormula = "{peminjaman.tgl_pinjam}>= cdate('" & DTPicker1.Value & "') AND {peminjaman.tgl_pinjam}<= CDATE('" & DTPicker2.Value & "')"
CrystalReport1.ReportFileName = App.Path & "\laporan_peminjaman.rpt"
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
ElseIf opt_cetakstruk.Value = True Then
CrystalReport1.SelectionFormula = "{peminjaman.no_peminjaman}='" & DataGrid1.Columns(0) & "'"
CrystalReport1.ReportFileName = App.Path + "\laporan_peminjaman.rpt"
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
Else: MsgBox "Pilih semua data ,salah satu atau data pertanggal yang akan dicetak", 0, "Pesan"
End If
End Sub


Private Sub Form_Load()
Call BukaDB
Dim koneksi As ADODB.Connection
Dim admin As ADODB.Recordset

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from peminjaman"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

 'pantullllllllllllllllllllllllllllllllllllllllllllllllllllll
 Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

pantul = 100
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub

Private Sub txt_cari_Change()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient

rs.Open "select * from peminjaman where no_peminjaman like '%" & txt_cari.Text & "%' or id_anggota like '%" & txt_cari.Text & "%' or tgl_pinjam like '%" & txt_cari.Text & "%' or tgl_kembali like '%" & txt_cari.Text & "%' or lama_sewa like '%" & txt_cari.Text & "%' or total_harga like '%" & txt_cari.Text & "%' or dibayar like '%" & txt_cari.Text & "%' or kembali like '%" & txt_cari.Text & "%' or total_pinjam like '%" & txt_cari.Text & "%' or keterangan_sudah_kembali like '%" & txt_cari.Text & "%'", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub
