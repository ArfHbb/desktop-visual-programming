VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daftar Pelanggan"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
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
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton Cbo_tutup 
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
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label lbl_sumber 
      Caption         =   "jgn hps"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form9"
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

Private Sub Cbo_tutup_Click()
If lbl_sumber.Caption = "pinjam" Then
Form3.Show
Form9.Hide
satupeminjam
Form3.txt_idfilm.SetFocus
ElseIf lbl_sumber.Caption = "kembali" Then
Form4.Show
Form9.Hide
End If
End Sub
Public Sub satupeminjam()
Form3.txt_idpinjam.Enabled = False
Form3.cbo_dialogpelanggan.Enabled = False
End Sub
Private Sub Form_Load()
 Dim l As Long
    l = GetWindowLong(Me.hwnd, GWL_STYLE)
    l = (l And Not WS_SYSMENU)
    l = SetWindowLong(Me.hwnd, GWL_STYLE, l)

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb;Persist Security Info=False"
Adodc1.RecordSource = "admin"
Adodc1.RecordSource = "select id_anggota,Nama, status from anggota where status='Member' or status='Non-member' order by id_Anggota"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
With DataGrid1
.Columns(0).Caption = "ID Pelanggan"
.Columns(0).Width = 1150

End With
DataGrid1.Refresh

Form10.Caption = "Dialog Pelanggan"

End Sub

Private Sub DataGrid1_DblClick()
If lbl_sumber.Caption = "pinjam" Then
Form3.txt_idpinjam.Text = DataGrid1.Columns(0)
Form9.Hide
satupeminjam
Form3.txt_idfilm.SetFocus

ElseIf lbl_sumber.Caption = "kembali" Then
Form4.txt_idanggota.Text = DataGrid1.Columns(0)
Form9.Hide
satupeminjam
End If
End Sub

Private Sub txt_cari_Change()
Set rs = New ADODB.Recordset
  If rs.State = adStateOpen Then
    rs.Close
    Set rs = New ADODB.Recordset
  End If
rs.CursorLocation = adUseClient
rs.Open "select id_anggota,Nama, status from anggota where id_Anggota like '%" & txt_cari.Text & "%' or Nama like '%" & txt_cari.Text & "%' and Status like '%" & txt_cari.Text & "%' and status='member' or 'non-member'", koneksi, adOpenStatic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Button = 2 Then
         PopupMenu MDIForm1.mnMenu, , X, Y
     End If
End Sub


