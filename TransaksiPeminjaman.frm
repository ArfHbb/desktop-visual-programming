VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form3 
   Caption         =   "ID Film"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   Icon            =   "TransaksiPeminjaman.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "TransaksiPeminjaman.frx":000C
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3. Pembayaran"
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
      Height          =   1815
      Left            =   1920
      TabIndex        =   19
      Top             =   5880
      Width           =   6615
      Begin VB.TextBox txt_dibayar 
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
         Left            =   4080
         TabIndex        =   23
         ToolTipText     =   "Bayar jumlah harga film"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga :"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah :"
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
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Uang Kembali :"
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
         Left            =   2760
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bayar Disini    :"
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
         Left            =   2760
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl_kembali 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lbl_total_harga 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   495
         Left            =   1080
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl_total_pinjam 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1080
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc DTDetailPinjam 
      Height          =   375
      Left            =   360
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "detail pinjam"
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
   Begin MSAdodcLib.Adodc DTPinjam 
      Height          =   330
      Left            =   360
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
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
      Caption         =   "peminjaman"
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
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1. Isi Transaksi Peminjaman"
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
      Height          =   4095
      Left            =   2040
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1440
      Width           =   6495
      Begin MSComCtl2.DTPicker date_pick_batas_kembali 
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   105054209
         CurrentDate     =   42715
      End
      Begin VB.CommandButton cbo_dialogpelanggan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Pilih ID Pelanggan"
         Top             =   1560
         Width           =   375
      End
      Begin MSComCtl2.DTPicker date_pick_pinjam 
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   2520
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   105054209
         CurrentDate     =   42702
      End
      Begin VB.TextBox txt_idpinjam 
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
         Left            =   3480
         TabIndex        =   7
         ToolTipText     =   "ID Pelanggan"
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Batas Pengembalian"
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
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Peminjaman"
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
         TabIndex        =   5
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Transaksi"
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
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lbl_nmrtrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Trans"
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
         Left            =   3480
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2. Data Peminjaman"
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
      Height          =   6255
      Left            =   8880
      TabIndex        =   8
      Top             =   1440
      Width           =   8055
      Begin VB.TextBox txt_idfilm 
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
         Left            =   3600
         TabIndex        =   25
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cbo_dialogidfilm 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
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
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Simpan Transaksi"
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
         TabIndex        =   18
         ToolTipText     =   "Pinjam !"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmd_batal 
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
         Left            =   3720
         MaskColor       =   &H00FF0000&
         TabIndex        =   17
         ToolTipText     =   "Batal meminjam"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H008080FF&
         Caption         =   "Keluar"
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
         TabIndex        =   11
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "&Pinjam"
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
         Left            =   -480
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4095
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   21
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
               LCID            =   1033
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
               LCID            =   1033
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
         Caption         =   "Cari Film yang akan dipinjam :"
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
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   7080
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
   Begin VB.Label txt_buatdelete 
      Caption         =   "jgn dihapus"
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi Peminjaman"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pantullllllllllllllllllllllllllllllllllllllllllllllll
Dim pantul As Integer


Private Sub cbo_dialogidfilm_Click()
Form11.Show
Form11.txt_cari.SetFocus
End Sub

Public Sub cbo_dialogpelanggan_Click()
Form9.Show
Form9.lbl_sumber.Caption = "pinjam"
End Sub

Private Sub cmd_cancel_Click()
Me.Hide
MDIForm1.Show
End Sub



Public Sub cmd_new_Click()
Call bersihwarna
Call BukaDB
Dim cari As New ADODB.Recordset
        cari.Open "select * from transaksi_pinjam where id_film='" & txt_idfilm & "'", koneksi
        cari.Requery
        
        If Not cari.EOF Then
                admin.Open "Select * from film where id_film ='" & txt_idfilm & "'", koneksi
               admin.Requery
               If Val(cari!jumlah + 1) > admin!Jumlah_stok Then
                MsgBox "Barang yang disewa melebihi stok yang ada", vbInformation, "Pesan"
                txt_idfilm.SetFocus
                txt_idfilm.BackColor = &H80FFFF
               Exit Sub
                End If
                If Not admin.EOF Then
                MsgBox "barang sudah dipesan", 0, "Pesan"
                txt_idfilm.SetFocus
                txt_idfilm.BackColor = &H80FFFF
                End If
        Else
            'jika id barang belum sama sekali dipilih
            koneksi.Close
            
            Call BukaDB
            admin.Open "Select * from film where id_film ='" & txt_idfilm & "'", koneksi
            admin.Requery
            If Not admin.EOF Then
                
                Adodc1.Recordset.MoveNext
                Adodc1.Recordset.MoveLast
                Adodc1.Recordset!id_film = admin!id_film
                Adodc1.Recordset!nama_film = admin!judul
                Adodc1.Recordset!jumlah = 1
                Adodc1.Recordset!Harga = 1 * admin!Harga
                Call Tambah_Baris
                Adodc1.Recordset.MoveNext
                Adodc1.Recordset.MoveLast
                Call JumlahHarga
                Call JumlahItem
                txt_idfilm.SetFocus
            Else
                MsgBox "Id film tidak terdaftar", vbInformation, "Pesan"
                txt_idfilm.SetFocus
                txt_idfilm.BackColor = &H80FFFF
                Exit Sub
            End If
        End If
        txt_idfilm.Text = ""
        txt_idfilm.SetFocus
        
End Sub


Function JumlahHarga()
Adodc1.Recordset.MoveFirst
A = 0

Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Harga <> vbNullString
    A = A + Adodc1.Recordset!Harga
    Adodc1.Recordset.MoveNext
    lbl_total_harga = Format(A, "###,###,###")
Loop
End Function

Function JumlahItem()
Adodc1.Recordset.MoveFirst
A = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!jumlah <> vbNullString
    A = A + Adodc1.Recordset!jumlah
    Adodc1.Recordset.MoveNext
    lbl_total_pinjam = A
Loop
End Function
Function Tambah_Baris()

For i = Adodc1.Recordset.RecordCount To Adodc1.Recordset.RecordCount
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!nomor = i + 1
    Adodc1.Recordset.update

Next i
End Function

Public Sub bersih()
txt_idpinjam.Text = ""
txt_idfilm.Text = ""
End Sub
Public Sub bersihwarna()
txt_dibayar.BackColor = &HFFFFFF
txt_idpinjam.BackColor = &HFFFFFF
txt_idfilm.BackColor = &HFFFFFF
End Sub
Private Sub Command2_Click()
'simpan ke tabel peminjaman
bersihwarna
If txt_dibayar.Text = "" Then
MsgBox "Peminjaman belum dibayar", 0, "Pesan"
txt_dibayar.SetFocus
txt_dibayar.BackColor = &H80FFFF
Else
DTPinjam.Recordset.AddNew
DTPinjam.Recordset!no_peminjaman = lbl_nmrtrans.Caption
DTPinjam.Recordset!id_anggota = txt_idpinjam.Text
DTPinjam.Recordset!tgl_pinjam = date_pick_pinjam
DTPinjam.Recordset!tgl_kembali = date_pick_batas_kembali
DTPinjam.Recordset!total_pinjam = lbl_total_pinjam.Caption
DTPinjam.Recordset!lama_sewa = 7
DTPinjam.Recordset!total_harga_sewa = Format(lbl_total_harga, "#########")
DTPinjam.Recordset!dibayar = Format(txt_dibayar, "#########")
DTPinjam.Recordset!keterangan_sudah_kembali = "belum"
If lbl_kembali = 0 Then
DTPinjam.Recordset!kembali = 0
Else
DTPinjam.Recordset!kembali = Format(lbl_kembali, "#########")
End If
DTPinjam.Recordset.update

'simpan ke tabel detail_peminjaman
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    If Adodc1.Recordset!id_film <> vbNullString Then
            DTDetailPinjam.Recordset.AddNew
            DTDetailPinjam.Recordset!no_peminjaman = lbl_nmrtrans.Caption
            DTDetailPinjam.Recordset!id_film = Adodc1.Recordset!id_film
            DTDetailPinjam.Recordset!jumlah = Adodc1.Recordset!jumlah
            DTDetailPinjam.Recordset!keterangan = "pinjam"
            DTDetailPinjam.Recordset.update
    End If
Adodc1.Recordset.MoveNext
Loop

'Pengurangan Jumlah Barang
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    If Adodc1.Recordset!id_film <> vbNullString Then
        Call BukaDB
        admin.Open "Select * from film where id_film='" & Adodc1.Recordset!id_film & "'", koneksi
        If Not admin.EOF Then
            Dim kurangi As String
            kurangi = "update film set jumlah_stok='" & admin!Jumlah_stok - Adodc1.Recordset!jumlah & "' where id_film='" & Adodc1.Recordset!id_film & "'"
            koneksi.Execute kurangi
        End If
    End If
Adodc1.Recordset.MoveNext
Loop
ini_no_transaksi = lbl_nmrtrans
Form_Activate
cmd_batal_Click

End If
txt_idpinjam.Enabled = True
cbo_dialogpelanggan.Enabled = True

If MsgBox("Transaksi peminjaman berhasil. Ingin mencetak Nota Peminjaman " + ini_no_transaksi + " ?", vbInformation + vbYesNo, "Pesan") = vbYes Then

cmd_batal_Click
cetak_struk

Else
cmd_batal_Click
End If


End Sub

Sub cetak_struk()
CrystalReport1.SelectionFormula = "{peminjaman.No_peminjaman}='" & ini_no_transaksi & "'"
CrystalReport1.ReportFileName = App.Path + "\laporan_peminjaman.rpt"
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End Sub

Private Sub cmd_batal_Click()
Tabel_Kosong
Unload Me
Form3.Show

End Sub

Private Sub Form_Activate()
bersihwarna
txt_idpinjam.Enabled = True
DataGrid1.Col = 1
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\rental.mdb"
    Adodc1.RecordSource = "transaksi_pinjam"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh
    If txt_idpinjam.Text = "" Then
    txt_idpinjam.SetFocus
    End If

DTPinjam.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\rental.mdb"
    DTPinjam.RecordSource = "peminjaman"
    DTPinjam.Refresh
    
DTDetailPinjam.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\rental.mdb"
    DTDetailPinjam.RecordSource = "detail_pinjam"
    DTDetailPinjam.Refresh
    
    
End Sub
Function Tabel_Kosong()
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
Loop
For i = 1 To 1
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!nomor = i
    Adodc1.Recordset.update
Next i
End Function


Private Sub date_pick_pinjam_Change()
date_pick_batas_kembali = DateAdd("d", 7, date_pick_pinjam.Value)
End Sub

Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from transaksi_pinjam"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'disini ada kodeotomatis
Call KodeOtomatis
Call bersih
Call bersihwarna

 'pantullllllllllllllllllllllllllllllllllllllllllllllllllllll
 Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

pantul = 100

date_pick_pinjam.Format = dtpCustom
date_pick_pinjam.CustomFormat = "dd MMMM yyyy"
date_pick_pinjam = Now
date_pick_batas_kembali.Format = dtpCustom
date_pick_batas_kembali.CustomFormat = "dd MMMM yyyy"
date_pick_batas_kembali = date_pick_pinjam + 7
End Sub

Sub KodeOtomatis()
Call BukaDB
admin.Open ("select * from peminjaman Where no_peminjaman In(Select Max(no_peminjaman)From peminjaman)Order By no_peminjaman Desc"), koneksi
admin.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With admin
        If .EOF Then
            Urutan = "PJM" + "001"
            lbl_nmrtrans = Urutan
        Else
            Hitung = Right(!no_peminjaman, 3) + 1
            Urutan = "PJM" + Right("000" & Hitung, 3)
        End If
        lbl_nmrtrans = Urutan
    End With
End Sub
Private Sub cari()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select * from transaksi_pinjam where Nomor like '%" & txt_cari.Text & "%' or id_film like '%" & txt_cari.Text & "%' or nama_film like '%" & txt_cari.Text & "%' or jumlah like '%" & txt_cari.Text & "%' or harga like '%" & txt_cari.Text & "%'", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

'pantullllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllll
Private Sub Timer1_Timer()
Label13.ForeColor = vbBlack
With Label13
 .Left = .Left + pantul
 If .Left < 0 Then pantul = 100
 If .Left > Me.ScaleWidth - .Width Then pantul = -100
 End With
End Sub
Private Sub txt_cari_Change()
Call cari
End Sub


Private Sub txt_dibayar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If txt_dibayar = "" Or Val(txt_dibayar) < (lbl_total_harga) Then
            MsgBox "Jumlah Pembayaran Kurang", vbInformation, "Pesan"
            lbl_kembali.Caption = ""
            txt_dibayar.SetFocus
            txt_dibayar.BackColor = &H80FFFF
        Else
            txt_dibayar = Format(txt_dibayar, "###,###,###")
            lbl_total_harga = Format(lbl_total_harga, "###,###,###")
            If txt_dibayar = lbl_total_harga Then
                lbl_kembali = Val(0)
            Else
                lbl_kembali = Format(txt_dibayar - lbl_total_harga, "###,###,###")
            End If
        Command2.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txt_idpinjam_Click()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from peminjaman"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub

