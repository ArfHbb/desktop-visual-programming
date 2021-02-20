VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "DataFIlm"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15705
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   15705
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5415
      Left            =   7920
      TabIndex        =   10
      Top             =   1200
      Width           =   7215
      Begin VB.CommandButton cmd_cancel4 
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
         Left            =   5400
         TabIndex        =   15
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmd_delete4 
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
         Left            =   3720
         TabIndex        =   14
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmd_save5 
         Caption         =   "&Save"
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
         Left            =   2040
         TabIndex        =   13
         Top             =   4080
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
         Left            =   360
         TabIndex        =   12
         Top             =   4080
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   360
         Top             =   3360
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2775
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox txt_iddisk 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txt_harga 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txt_judul 
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txt_idfilm 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "ID Disk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Harga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Judul"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID Film"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   870
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DATA FILM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select * from film"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub cmd_delete_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbCritical + vbOKCancel, "Hati-hati menghapus data!!") = vbOK Then
Adodc1.Recordset.Delete
MsgBox "Data Berhasil Di Hapus", vbInformation + vbOKOnly, "Pesan"
DataGrid1.Refresh
End If
End Sub

Private Sub cmd_cancel_Click()
Call Cancel
End Sub

Private Sub cmd_save_Click()
Adodc1.Recordset.Fields("id_film") = txt_idfilm.Text
Adodc1.Recordset.Fields("Judul") = txt_judul.Text
Adodc1.Recordset.Fields("Harga") = txt_harga.Text
Adodc1.Recordset.Fields("id_disk") = txt_iddisk.Text
Adodc1.Recordset.Update
MsgBox "Data Berhasil Di Edit", vbInformation + vbOKOnly, "Pesan"
End Sub

Private Sub cmd_new_Click()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from pengembalian"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("id_film") = txt_idfilm.Text
Adodc1.Recordset.Fields("Judul") = txt_judul.Text
Adodc1.Recordset.Fields("Harga") = txt_harga.Text
Adodc1.Recordset.Fields("id_disk") = txt_iddisk.Text
Adodc1.Recordset.Update
End Sub

Private Sub DataGrid1_Click()
txt_idfilm.Text = Adodc1.Recordset!id_film
txt_judul.Text = Adodc1.Recordset!Judul
txt_harga.Text = Adodc1.Recordset!Harga
txt_iddisk.Text = Adodc1.Recordset!id_disk
End Sub
