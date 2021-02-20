VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "ID Film"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1320
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6255
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1080
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   6720
         TabIndex        =   22
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox txt_tglkembl 
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox txt_tlp 
         Enabled         =   0   'False
         Height          =   405
         Left            =   4080
         TabIndex        =   14
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox txt_almtpmnjm 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txt_nmpmnjm 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txt_idpinjam 
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "Tanggal Pengembalian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label11 
         Caption         =   "Tanggal Peminjaman"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Telp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "ID Pelanggan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "No. Transaksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lbl_tglpnjm 
         Caption         =   "Tgl Peminjaman"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label lbl_nmrtrans 
         Caption         =   "No. Trans"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6255
      Left            =   7680
      TabIndex        =   16
      Top             =   1080
      Width           =   8055
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   21
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton btn_hapus 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   20
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   19
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   5160
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3735
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6588
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   -120
      Top             =   7560
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
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
      Caption         =   "TRANSAKSI PEMINJAMAN FILM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_hapus_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbCritical + vbOKCancel, "Hati-hati menghapus data!!") = vbOK Then
Adodc1.Recordset.Delete
DataGrid1.Refresh
End If
End Sub

Private Sub cmd_cancel_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub cmd_edit_Click()
Adodc1.Recordset.Fields("No_peminjaman") = lbl_nmrtrans.Caption
Adodc1.Recordset.Fields("id_anggota") = txt_idpinjam.Text
Adodc1.Recordset.Fields("Tgl_pinjam") = txt_tglpnjm.Caption
Adodc1.Recordset.Update
MsgBox "Data Berhasil Di Edit", vbInformation + vbOKOnly, "Pesan"
End Sub

Private Sub cmd_new_Click()
Call BukaDB
Call KodeOtomatis
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from peminjaman"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("No_peminjaman") = lbl_nmrtrans.Caption
Adodc1.Recordset.Fields("id_anggota") = txt_idpinjam.Text
Adodc1.Recordset.Fields("Tgl_pinjam") = lbl_tglpnjm.Caption

Adodc1.Recordset.Update
End Sub


Private Sub Command1_Click()
Call BukaDB2
Admin2.Open "Select * from anggota where id_anggota ='" & txt_idpinjam & "'", koneksi
        If Admin2.EOF Then
            MsgBox "Datatidakada!"
        Else
        txt_nmpmnjm.Text = Adodc2.Recordset!nama
        txt_almtpmnjm.Text = Adodc2.Recordset!alamat
        txt_tlp.Text = Adodc2.Recordset!telephone
        End If

    End Sub

Private Sub DataGrid1_Click()
lbl_nmrtrans.Caption = Adodc1.Recordset!no_peminjaman
txt_idpinjam.Text = Adodc1.Recordset!id_anggota
lbl_tglpnjm.Caption = Adodc1.Recordset!Tgl_pinjam
txt_idpinjam.Text = Adodc1.Recordset!id_anggota
     
End Sub

Private Sub Form_Load()
lbl_tglpnjm = Date
Call BukaDB
Call BukaDB2
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from peminjaman"
Adodc1.Refresh
'mulai sini
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc2.RecordSource = "admin"
Adodc2.RecordSource = "select * from anggota"
Adodc2.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'disini ada kodeotomatis
Call KodeOtomatis
End Sub

Sub KodeOtomatis()
Call BukaDB
Admin.Open ("select * from peminjaman Where no_peminjaman In(Select Max(no_peminjaman)From peminjaman)Order By no_peminjaman Desc"), koneksi
Admin.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With Admin
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
