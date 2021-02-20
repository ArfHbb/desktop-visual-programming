VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "dialog_peminjaman"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1080
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.CommandButton tutup 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4471
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
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
Attribute VB_Name = "Form10"
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

Public Sub DataGrid1_DblClick()
DataGrid1.Refresh
Form4.txt_nopinjam.Text = DataGrid1.Columns(0)
Form4.txt_idanggota.Text = DataGrid1.Columns(1)
Form4.date_pick_pinjam = DataGrid1.Columns(2)
Form10.Hide
Form4.tampilkanfilmyangdipinjamkedatagrid

tutup_Click
End Sub

Private Sub txt_cari_Change()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select no_peminjaman,id_anggota, tgl_pinjam from peminjaman where keterangan_sudah_kembali='belum' and No_peminjaman like '%" & txt_cari.Text & "%' or id_anggota like '%" & txt_cari.Text & "%' or tgl_kembali like '%" & txt_cari.Text & "%' or tgl_pinjam like '%" & txt_cari.Text & "%' order by id_Anggota", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

Private Sub tutup_Click()
Form10.Hide
Form4.Hide
Unload Me
Form4.Show
End Sub

Private Sub Form_Load()
 Dim l As Long
    l = GetWindowLong(Me.hwnd, GWL_STYLE)
    l = (l And Not WS_SYSMENU)
    l = SetWindowLong(Me.hwnd, GWL_STYLE, l)

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select no_peminjaman,id_anggota, tgl_pinjam from peminjaman where keterangan_sudah_kembali='belum' order by id_Anggota"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
With DataGrid1
.Columns(0).Caption = "No Peminjaman"
.Columns(1).Caption = "ID Pelanggan"
.Columns(2).Caption = "Tanggal Pinjam"

End With
DataGrid1.Refresh
Form10.Caption = "Dialog Peminjaman"

End Sub

