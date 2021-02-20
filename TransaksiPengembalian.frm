VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14970
   Icon            =   "TransaksiPengembalian.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "TransaksiPengembalian.frx":000C
   ScaleHeight     =   8430
   ScaleWidth      =   14970
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "4. Tanggal Pinjam Dan Kembali"
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
      Height          =   1215
      Left            =   3120
      TabIndex        =   28
      ToolTipText     =   "Pilih tanggal pinjam dan kembali"
      Top             =   3960
      Width           =   6255
      Begin MSComCtl2.DTPicker date_pick_pinjam 
         Height          =   375
         Left            =   3600
         TabIndex        =   30
         ToolTipText     =   "Tanggal Peminjaman Film"
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
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
         CalendarForeColor=   16711680
         CalendarTitleForeColor=   16711680
         Format          =   51838977
         CurrentDate     =   42708
      End
      Begin MSComCtl2.DTPicker date_pick_kembali 
         Height          =   375
         Left            =   3600
         TabIndex        =   32
         ToolTipText     =   "Tanggal Pengembalian Film"
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
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
         CalendarForeColor=   16711680
         CalendarTitleForeColor=   16711680
         Format          =   51838977
         CurrentDate     =   42708
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Kembali"
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
         Height          =   360
         Left            =   0
         TabIndex        =   31
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal  Pinjam"
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
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2. Pilih film yang akan dikembalikan :"
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
      Height          =   1935
      Left            =   9720
      TabIndex        =   25
      Top             =   960
      Width           =   8775
      Begin MSDataGridLib.DataGrid DGTelahPinjam 
         Height          =   1575
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "Dobel Klik Data film yang telah dipinjam"
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2778
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "5. Pembayaran Denda"
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
      Height          =   3015
      Left            =   3120
      TabIndex        =   16
      ToolTipText     =   "Bayar denda jika ada"
      Top             =   5280
      Width           =   6255
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
         Left            =   3120
         TabIndex        =   23
         ToolTipText     =   "Bayar denda"
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txt_denda 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   3120
         TabIndex        =   18
         ToolTipText     =   "Jumlah Denda yang harus dibayar"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bayar Disini :"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Film yang dipinjam :"
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
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pelanggan :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl_nama_pelanggan 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
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
         Left            =   3120
         TabIndex        =   20
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lbl_jumlah 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
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
         Left            =   3120
         TabIndex        =   19
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Denda (Rp)"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc DTCariTelahPinjam 
      Height          =   330
      Left            =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc DTDetailKembali 
      Height          =   375
      Left            =   0
      Top             =   6000
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc DTKembali 
      Height          =   375
      Left            =   0
      Top             =   5400
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
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3. Data Pengembalian"
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
      Height          =   5175
      Left            =   9720
      TabIndex        =   3
      Top             =   3120
      Width           =   8775
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
         Left            =   3240
         MaskColor       =   &H00FF0000&
         TabIndex        =   15
         ToolTipText     =   "Batal mengembalikan CD/DVD"
         Top             =   4440
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3855
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Data Film yang telah dikembalikan"
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6800
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         Left            =   6360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         BackColor       =   &H00FF8080&
         Caption         =   "Simpan"
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
         TabIndex        =   4
         ToolTipText     =   "Kembalikan CD/DVD"
         Top             =   4440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1. Isi Transaksi Pengembalian"
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
      Height          =   2655
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   6255
      Begin VB.CommandButton cbo_dialogpeminjaman 
         Appearance      =   0  'Flat
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
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Pilih peminjaman"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cbo_dialogpelanggan 
         Appearance      =   0  'Flat
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txt_idanggota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   3360
         TabIndex        =   10
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txt_nopinjam 
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
         Left            =   3360
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lbl_nokmbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "no pengembalian"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No Pengembalian"
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
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No Peminjaman"
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
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Label lbl_kembali 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   255
      Left            =   -360
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label txt_buatdelete 
      Caption         =   "jgn dihaus"
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi Pengembalian"
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
      Left            =   3600
      TabIndex        =   9
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumlah_film As Integer
'pantullllllllllllllllllllllllllllllllllllllllllllllll
Dim pantul As Integer
Public Sub bersih()
txt_nopinjam.Text = ""
txt_idanggota.Text = ""
txt_denda.Text = ""
End Sub

Private Sub cmd_batal_Click()
Tabel_Kosong
Unload Me
Form4.Show
End Sub

Public Sub DGTelahPinjam_DblClick()
    Call BukaDB
    Dim cari As New ADODB.Recordset
    cari.Open "select * from transaksi_kembali where nomor_peminjaman= '" & DTCariTelahPinjam.Recordset!no_peminjaman & "' and id_film= '" & DTCariTelahPinjam.Recordset!id_film & "'", koneksi
    If Not cari.EOF Then
        MsgBox "data jangan dientri dua kali", vbInformation, "Pesan"
        Exit Sub
    Else
        Call SelectAllVisible
    End If
    
    Dim selisih, dendahari As Integer
 selisih = DateDiff("d", date_pick_pinjam.Value, date_pick_kembali.Value)
 If selisih > 7 Then
 dendahari = ((selisih - 7) * 1000)
 txt_denda.Text = dendahari * Val(lbl_jumlah)
 ElseIf selisih <= 7 Then
 Call denda0
 ElseIf selisih = 0 Then
 Call denda0
 

 End If
  txt_dibayar.SetFocus
Dim denda As Integer
denda = txt_denda.Text
If denda < 0 Then
MsgBox "Maaf tanggal yang dimasukkan salah !", vbInformation, "Informasi"
txt_denda.Text = ""
date_pick_pinjam.SetFocus
End If
txt_denda = Format(txt_denda, "###,###,###")
End Sub

Private Sub Form_Load()


 'pantullllllllllllllllllllllllllllllllllllllllllllllllllllll
 Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2

pantul = 100
jml_film = 0
lbl_jumlah = jml_film
lbl_jumlah.Visible = True
txt_denda.MaxLength = 6
date_pick_pinjam.Format = dtpCustom
date_pick_pinjam.CustomFormat = "dd MMMM yyyy"
date_pick_kembali.Format = dtpCustom
date_pick_kembali.CustomFormat = "dd MMMM yyyy"
date_pick_kembali = Now
lbl_kbl = Now
End Sub
Private Sub Form_Activate()

Call BukaDB
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\rental.mdb"
    Adodc1.RecordSource = "transaksi_kembali"
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh
    Call KodeOtomatis
   
    Adodc1.Recordset.MoveFirst
    DataGrid1.Col = 1
    
DTKembali.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\rental.mdb"
    DTKembali.RecordSource = "pengembalian"
    DTKembali.Refresh
    
DTDetailKembali.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\rental.mdb"
    DTDetailKembali.RecordSource = "detail_kembali"
    DTDetailKembali.Refresh
    
lbl_nama_pelanggan.Caption = ""
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

Public Sub DGTelahPinjam_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
    Call BukaDB
    Dim cari As New ADODB.Recordset
    cari.Open "select * from transaksi_kembali where nomor_peminjaman= '" & DTCariTelahPinjam.Recordset!no_peminjaman & "'", koneksi
    If Not cari.EOF Then
        MsgBox "data jangan dientri dua kali", vbInformation, "Pesan"
        Exit Sub
    Else
        Call SelectAllVisible
    End If
    
End Select
End Sub

Sub SelectAllVisible()
'On Error Resume Next
    Adodc1.Recordset!nomor_peminjaman = DGTelahPinjam.Columns(0)
    Adodc1.Recordset!id_film = DGTelahPinjam.Columns(1)
    Adodc1.Recordset!judul = DGTelahPinjam.Columns(2)
    Adodc1.Recordset!tanggal = DGTelahPinjam.Columns(3)
    Adodc1.Recordset!jumlah = DGTelahPinjam.Columns(5)
    Adodc1.Recordset!denda = (CDate(date_pick_kembali) - (Adodc1.Recordset!tanggal) - 3) * 500 * Adodc1.Recordset!jumlah
    
    If Adodc1.Recordset!denda < 0 Then
        Adodc1.Recordset!denda = 0
    End If

    Call Tambah_Baris
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 1
    Adodc1.Recordset.MoveLast
    Call TotalKbl
    Call JmlDenda
    Call JumlahItem

End Sub

Function JumlahItem()
Adodc1.Recordset.MoveFirst
A = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!jumlah <> vbNullString
    A = A + Adodc1.Recordset!jumlah
    Adodc1.Recordset.MoveNext
    lbl_jumlah = A
Loop
End Function

Function Tambah_Baris()
For i = Adodc1.Recordset.RecordCount To Adodc1.Recordset.RecordCount
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!nomor = i + 1
    Adodc1.Recordset.update
Next i
End Function

Function TotalKbl()
Adodc1.Recordset.MoveFirst
A = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!jumlah <> vbNullString
    A = A + Adodc1.Recordset!jumlah
    Adodc1.Recordset.MoveNext
    lbl_total_kembali = A
Loop
End Function

Function JmlDenda()
Adodc1.Recordset.MoveFirst
A = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!nomor_peminjaman <> vbNullString
    A = A + Adodc1.Recordset!denda
    Adodc1.Recordset.MoveNext
    lbl_total_denda = A
Loop
End Function

Private Sub txt_dibayar_KeyPress(KeyAscii As Integer)
'if denda 0 eror bro!
If KeyAscii = 13 Then
        If txt_dibayar = "" Or Val(txt_dibayar) < Val(txt_denda) Then
            MsgBox "Jumlah Pembayaran Kurang", vbInformation, "Pesan"
            txt_dibayar.SetFocus
        Else
            txt_dibayar = Format(txt_dibayar, "###,###,###")
            If txt_dibayar = lbl_total_denda Then
                lbl_kembali = 0
            ElseIf txt_denda = 0 Then
                txt_dibayar = txt_denda
            Else
                lbl_kembali = Format(txt_dibayar - lbl_total_denda, "###,###,###")
            End If
       
        End If
        cmd_new.SetFocus
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub


Function Telah_Pinjam()
    On Error Resume Next
    Set TP = New ADODB.Recordset
    TP.Open "SELECT sum(total_pinjam) AS JUMTOTAL FROM peminjaman WHERE id_anggota='" & txt_idanggota & "'", koneksi
    Telah_Pinjam = TP!JumTotal
    If Telah_Pinjam > 0 Then
    lbl_telah_pinjam = Telah_Pinjam
    Else
    lbl_telah_pinjam = 0
    End If
End Function

Sub Pinjaman()
    DTCariTelahPinjam.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\rental.mdb"
    DTCariTelahPinjam.RecordSource = "Select Distinct detail_pinjam.no_peminjaman,film.id_film,film.judul,peminjaman.tgl_pinjam,peminjaman.tgl_kembali,detail_pinjam.jumlah,(Date()-peminjaman.tgl_pinjam)+1 as [LamaPinjam] From  anggota,film,peminjaman,detail_pinjam Where film.id_film=detail_pinjam.id_film And peminjaman.no_peminjaman=Left(detail_pinjam.no_peminjaman,6) And anggota.id_anggota=peminjaman.id_anggota and keterangan='pinjam' and detail_pinjam.no_peminjaman ='" & txt_nopinjam & "' And anggota.id_anggota='" & txt_idanggota & "'"

    DTCariTelahPinjam.Refresh
    Set DGTelahPinjam.DataSource = DTCariTelahPinjam
    DGTelahPinjam.Refresh
End Sub

Public Sub txt_idanggota_KeyPress(KeyAscii As Integer)
Call tampilkanfilmyangdipinjamkedatagrid
End Sub

Public Sub tampilkanfilmyangdipinjamkedatagrid()
BukaDB
BukaDB

    admin.Open "Select * from anggota where id_anggota='" & txt_idanggota & "'", koneksi
    If Not admin.EOF Then
        lbl_nama_pelanggan.Caption = admin!nama
    Else
        MsgBox "Id pelanggan tidak terdaftar", vbInformation, "Pesan"
        Exit Sub
    End If
    
    'koneksi untuk cari data yang dipinjam pelanggan
    Call Telah_Pinjam
    Call Pinjaman
    
    If Telah_Pinjam = 0 Or lbl_telah_pinjam = "" Or lbl_telah_pinjam = 0 Then
        'MsgBox "" & lbl_nama_pelanggan & " tidak punya pinjaman", vbInformation, "Pesan"
        'txt_idanggota.SetFocus
        'disini masih eror
        Exit Sub
    Else
    End If
    
End Sub

Private Sub cmd_cancel_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub cmd_new_Click()
If lbl_jumlah = 0 Then
MsgBox "Tidak ada film yang akan dikembalikan", 0, "Pesan"
ElseIf txt_dibayar.Text = "" Then
MsgBox "Denda Belom dibaya", 0, "Pesan"
ElseIf lbl_jumlah = 0 Then
MsgBox "Pilih Film untuk dikembalikan", 0, "Pesan"
Else
BukaDB
    admin.Open "Select * from anggota where id_anggota='" & txt_idanggota & "'", koneksi
    If Not admin.EOF Then
        'If lbl_telah_pinjam <= 0 Then
        'MsgBox RSRental!nama_pelanggan + " belum melakukan peminjaman " + vbNewLine, vbInformation, "Pesan"
        'Exit Sub
        'Else
        'End If
    Else
        MsgBox "Id pelanggan tidak terdaftar", vbInformation, "Pesan"
        txt_idanggota.SetFocus
        Exit Sub
    End If

'simpan ke tabel pengembalian
DTKembali.Recordset.AddNew
DTKembali.Recordset!no_pengembalian = lbl_nokmbl.Caption
DTKembali.Recordset!id_anggota = txt_idanggota.Text
DTKembali.Recordset!tgl_kembali = date_pick_kembali
If txt_denda.Text = "" Then
DTKembali.Recordset!denda = 0
Else
DTKembali.Recordset!denda = txt_denda
End If
DTKembali.Recordset!dibayar = txt_dibayar
DTKembali.Recordset!kembali = lbl_kembali
DTKembali.Recordset.update

'simpan ke tabel detail_pengembalian
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    If Adodc1.Recordset!id_film <> vbNullString Then
            DTDetailKembali.Recordset.AddNew
            DTDetailKembali.Recordset!nomor_pengembalian = lbl_nokmbl.Caption
            DTDetailKembali.Recordset!id_film = Adodc1.Recordset!id_film
            DTDetailKembali.Recordset!jumlah_film = Adodc1.Recordset!jumlah
            DTDetailKembali.Recordset.update
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
            Dim tambah As String
            tambah = "update film set jumlah_stok='" & admin!Jumlah_stok + Adodc1.Recordset!jumlah & "' where id_film='" & Adodc1.Recordset!id_film & "'"
            koneksi.Execute tambah
        End If
    End If
Adodc1.Recordset.MoveNext
Loop

'update keterangan pinjaman
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    If Adodc1.Recordset!nomor_peminjaman <> vbNullString Then
        Call BukaDB
        admin.Open "Select * from detail_pinjam where no_peminjaman='" & Adodc1.Recordset!nomor_peminjaman & "' and keterangan='pinjam'", koneksi
        If Not admin.EOF Then
            Dim update As String
            update = "update detail_pinjam set keterangan='kembali' where no_peminjaman ='" & Adodc1.Recordset!nomor_peminjaman & "'"
            koneksi.Execute (update)
            
        End If
    End If
Adodc1.Recordset.MoveNext
Loop


'tampilkan yang hanya belum dikembalikan
Call BukaDB
        admin.Open "Select * from peminjaman where No_peminjaman='" & txt_nopinjam & "'", koneksi
        If Not admin.EOF Then
        Dim updatekembali As String
            updatekembali = "update peminjaman set Keterangan_sudah_kembali='sudah' where No_peminjaman ='" & txt_nopinjam & "'"
            koneksi.Execute (updatekembali)
        End If
   

Form_Activate
cmd_batal_Click
End If

End Sub

Public Sub baru()
Call BukaDB
' ini codng mencegah data redundant
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select*from pengembalian where no_pengembalian='" & lbl_nokmbl.Caption & "'", koneksi

If rs.EOF Then 'jika tidak ditemukan, maka isi data
Call BukaDB
Call KodeOtomatis
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from pengembalian"
'disinii

Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("No_pengembalian") = lbl_nokmbl.Caption
Adodc1.Recordset.Fields("No_peminjaman") = txt_nopinjam.Text
Adodc1.Recordset.Fields("id_anggota") = txt_idanggota.Text
Adodc1.Recordset.Fields("Tgl_pinjam") = date_pick_pinjam
Adodc1.Recordset.Fields("Tgl_kembali") = date_pick_kembali
Adodc1.Recordset.Fields("denda") = txt_denda.Text
Adodc1.Recordset.update

Else 'jika ditemukan, maka ada pesan data sudah ada
A = rs!no_pengembalian
MsgBox "Data Dengan no_pengembalian " & A & " Sudah ADA", vbCritical, "SIMPAN"
End If 'sampek sisni pencegahan redundant

End Sub

Public Sub cbo_dialogpelanggan_Click()
Form9.Show
Form9.lbl_sumber.Caption = "kembali"
Form9.txt_cari.SetFocus
End Sub

Public Sub Cbo_dialogpeminjaman_Click()
Form4.Hide
Form4.Show
Form10.Hide
Form10.Show

Form10.txt_cari.SetFocus
End Sub

Public Sub denda0()
txt_dibayar.Text = "0"
 txt_denda.Text = "0"
 lbl_kembali.Caption = "0"
End Sub



Sub KodeOtomatis()
Call BukaDB
admin.Open ("select * from pengembalian Where no_pengembalian In(Select Max(no_pengembalian)From pengembalian)Order By no_pengembalian Desc"), koneksi
admin.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With admin
        If .EOF Then
            Urutan = "KBL" + "001"
            lbl_nokmbl = Urutan
        Else
            Hitung = Right(!no_pengembalian, 3) + 1
            Urutan = "KBL" + Right("000" & Hitung, 3)
        End If
        lbl_nokmbl = Urutan
    End With
End Sub
Private Sub cari()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select * from transaksi_kembali where nomor like '%" & txt_cari.Text & "%' or nomor_peminjaman like '%" & txt_cari.Text & "%' or id_film like '%" & txt_cari.Text & "%' or tanggal like '%" & txt_cari.Text & "%' or judul like '%" & txt_cari.Text & "%' or jumlah like '%" & txt_cari.Text & "%' or denda like '%" & txt_cari.Text & "%'", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub
Private Sub txt_cari_Change()
Call cari
End Sub

Private Sub txt_denda_KeyPress(KeyAscii As Integer)
'tidak bisa diidsi angka atau huruf
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub

