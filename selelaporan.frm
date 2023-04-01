VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form selelaporan 
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox sisaaa 
      Height          =   375
      Left            =   9240
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1980
      Width           =   915
   End
   Begin VB.Timer sementarat 
      Interval        =   1
      Left            =   10470
      Top             =   1560
   End
   Begin VB.Timer picket2 
      Interval        =   1
      Left            =   12360
      Top             =   5910
   End
   Begin VB.TextBox satuan 
      Height          =   375
      Left            =   2790
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   1020
      Width           =   2175
   End
   Begin VB.TextBox midoftgl 
      Height          =   405
      Left            =   6720
      TabIndex        =   27
      Top             =   1590
      Width           =   825
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6780
      Top             =   690
   End
   Begin VB.TextBox noww 
      Height          =   405
      Left            =   390
      TabIndex        =   26
      Text            =   "15/01/2019"
      Top             =   120
      Width           =   2025
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   330
      Left            =   2520
      TabIndex        =   25
      Top             =   330
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   582
      _Version        =   393216
      Format          =   103022593
      CurrentDate     =   43814
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TES"
      Height          =   285
      Left            =   11160
      TabIndex        =   24
      Top             =   3510
      Width           =   495
   End
   Begin VB.TextBox outstok3 
      Height          =   405
      Left            =   11130
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   480
      Width           =   1965
   End
   Begin VB.TextBox sisaa3 
      Height          =   405
      Left            =   8130
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   510
      Width           =   1965
   End
   Begin VB.TextBox outstok2 
      Height          =   405
      Left            =   11100
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1050
      Width           =   1965
   End
   Begin VB.TextBox sisaa2 
      Height          =   405
      Left            =   8100
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1080
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TES"
      Height          =   285
      Left            =   11580
      TabIndex        =   18
      Top             =   5580
      Width           =   495
   End
   Begin VB.Timer picket 
      Interval        =   1
      Left            =   12360
      Top             =   4500
   End
   Begin VB.Timer deletet 
      Interval        =   10
      Left            =   12360
      Top             =   3960
   End
   Begin VB.TextBox idob 
      Height          =   375
      Left            =   390
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1050
      Width           =   2175
   End
   Begin VB.Timer updatet 
      Interval        =   10
      Left            =   12360
      Top             =   3420
   End
   Begin VB.Timer savet 
      Interval        =   10
      Left            =   12360
      Top             =   2850
   End
   Begin VB.TextBox outstok 
      Height          =   375
      Left            =   11100
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1980
      Width           =   2175
   End
   Begin VB.TextBox sisaa 
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2010
      Width           =   1155
   End
   Begin VB.TextBox tanggalaa 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2010
      Width           =   2175
   End
   Begin VB.TextBox jumlahaa 
      Height          =   375
      Left            =   2850
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2010
      Width           =   2175
   End
   Begin VB.TextBox namaa 
      Height          =   375
      Left            =   390
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2010
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid dgpen 
      Height          =   4515
      Left            =   330
      TabIndex        =   0
      Top             =   2520
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7964
      _Version        =   393216
      BackColor       =   16384
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   22
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
         Name            =   "Trebuchet MS"
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
   Begin MSAdodcLib.Adodc adopen 
      Height          =   330
      Left            =   12090
      Top             =   4980
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\SMK\VISUAL BASIC\P U S K E S M A S\USKES.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\SMK\VISUAL BASIC\P U S K E S M A S\USKES.mdb;Persist Security Info=False"
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
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   375
      Left            =   11220
      TabIndex        =   32
      Top             =   6840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label13 
      Caption         =   "Selector(Tambah Stok)"
      Height          =   465
      Left            =   11070
      TabIndex        =   30
      Top             =   5940
      Width           =   1395
   End
   Begin VB.Label Label12 
      Caption         =   "Satuan"
      Height          =   405
      Left            =   2790
      TabIndex        =   29
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label11 
      Caption         =   "FRM Jumlah"
      Height          =   405
      Left            =   10080
      TabIndex        =   21
      Top             =   1110
      Width           =   915
   End
   Begin VB.Label stts 
      Caption         =   "ADA"
      Height          =   405
      Left            =   11160
      TabIndex        =   17
      Top             =   2430
      Width           =   915
   End
   Begin VB.Label Label10 
      Caption         =   "Selector"
      Height          =   195
      Left            =   11670
      TabIndex        =   16
      Top             =   4590
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "Delete"
      Height          =   195
      Left            =   11670
      TabIndex        =   15
      Top             =   4050
      Width           =   915
   End
   Begin VB.Label Label8 
      Caption         =   "Update"
      Height          =   195
      Left            =   11670
      TabIndex        =   14
      Top             =   3510
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "Save"
      Height          =   195
      Left            =   11670
      TabIndex        =   13
      Top             =   2940
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "IDObat"
      Height          =   405
      Left            =   390
      TabIndex        =   12
      Top             =   780
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Stok yang tersisa"
      Height          =   405
      Left            =   8100
      TabIndex        =   10
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Total Pengeluaran"
      Height          =   405
      Left            =   11190
      TabIndex        =   8
      Top             =   1530
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal"
      Height          =   405
      Left            =   5370
      TabIndex        =   7
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah Stok Awal Masuk"
      Height          =   405
      Left            =   2880
      TabIndex        =   6
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      Height          =   405
      Left            =   420
      TabIndex        =   5
      Top             =   1530
      Width           =   915
   End
End
Attribute VB_Name = "selelaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
picket.Enabled = True
End Sub

Private Sub Command2_Click()
updatet.Enabled = True
End Sub

Private Sub deletet_Timer()
idob.Text = ""
namaa.Text = ""
jumlahaa.Text = ""
tanggalaa.Text = ""
outstok.Text = ""
outstok2.Text = ""
outstok3.Text = ""
sisaa.Text = ""
sisaa2.Text = ""
sisaa3.Text = ""
midoftgl.Text = ""
adopen.RecordSource = "LPRSTOK"
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.AllowUpdate = False
dgpen.TabStop = False
dgpen.Refresh
deletet.Enabled = False
End Sub

Private Sub Form_Load()
sql = "select * from LPRSTOK"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
savet.Enabled = False
updatet.Enabled = False
picket.Enabled = False
deletet.Enabled = False
picket2.Enabled = False
sementarat.Enabled = False
Label21.Caption = adopen.Recordset.RecordCount
End Sub

Private Sub picket_Timer()
If resep2.index.Text = "1" Then
sql = "select * from LPRSTOK where NMGenerik like '%" & resep2.nb1.Text & "%' order by NMGenerik asc"
adopen.RecordSource = sql
adopen.Refresh
idob.Text = adopen.Recordset.Fields(0)
namaa.Text = adopen.Recordset.Fields(1)
satuan.Text = adopen.Recordset.Fields(2)
jumlahaa.Text = adopen.Recordset.Fields(3)
tanggalaa.Text = adopen.Recordset.Fields(4)
sisaa.Text = adopen.Recordset.Fields(5)
sisaaa.Text = adopen.Recordset.Fields(5)
outstok.Text = adopen.Recordset.Fields(6)
sisaa3.Text = Val(sisaa.Text) - Val(sisaa2.Text)
outstok3.Text = Val(outstok.Text) + Val(outstok2.Text)
sisaa.Text = sisaa3.Text
outstok.Text = outstok3.Text
midoftgl.Text = Mid(tanggalaa.Text, 4, 2)
picket.Enabled = False
updatet.Enabled = True
ElseIf resep2.index.Text = "2" Then
sql = "select * from LPRSTOK where NMGenerik like '%" & resep2.nb2.Text & "%' order by NMGenerik asc"
adopen.RecordSource = sql
adopen.Refresh
idob.Text = adopen.Recordset.Fields(0)
namaa.Text = adopen.Recordset.Fields(1)
satuan.Text = adopen.Recordset.Fields(2)
jumlahaa.Text = adopen.Recordset.Fields(3)
tanggalaa.Text = adopen.Recordset.Fields(4)
sisaa.Text = adopen.Recordset.Fields(5)
sisaaa.Text = adopen.Recordset.Fields(5)
outstok.Text = adopen.Recordset.Fields(6)
sisaa3.Text = Val(sisaa.Text) - Val(sisaa2.Text)
outstok3.Text = Val(outstok.Text) + Val(outstok2.Text)
sisaa.Text = sisaa3.Text
outstok.Text = outstok3.Text
midoftgl.Text = Mid(tanggalaa.Text, 4, 2)
picket.Enabled = False
updatet.Enabled = True
ElseIf resep2.index.Text = "3" Then
sql = "select * from LPRSTOK where NMGenerik like '%" & resep2.nb3.Text & "%' order by NMGenerik asc"
adopen.RecordSource = sql
adopen.Refresh
idob.Text = adopen.Recordset.Fields(0)
namaa.Text = adopen.Recordset.Fields(1)
satuan.Text = adopen.Recordset.Fields(2)
jumlahaa.Text = adopen.Recordset.Fields(3)
tanggalaa.Text = adopen.Recordset.Fields(4)
sisaa.Text = adopen.Recordset.Fields(5)
sisaaa.Text = adopen.Recordset.Fields(5)
outstok.Text = adopen.Recordset.Fields(6)
sisaa3.Text = Val(sisaa.Text) - Val(sisaa2.Text)
outstok3.Text = Val(outstok.Text) + Val(outstok2.Text)
sisaa.Text = sisaa3.Text
outstok.Text = outstok3.Text
midoftgl.Text = Mid(tanggalaa.Text, 4, 2)
picket.Enabled = False
updatet.Enabled = True
ElseIf resep2.index.Text = "4" Then
sql = "select * from LPRSTOK where NMGenerik like '%" & resep2.nb4.Text & "%' order by NMGenerik asc"
adopen.RecordSource = sql
adopen.Refresh
idob.Text = adopen.Recordset.Fields(0)
namaa.Text = adopen.Recordset.Fields(1)
satuan.Text = adopen.Recordset.Fields(2)
jumlahaa.Text = adopen.Recordset.Fields(3)
tanggalaa.Text = adopen.Recordset.Fields(4)
sisaa.Text = adopen.Recordset.Fields(5)
sisaaa.Text = adopen.Recordset.Fields(5)
outstok.Text = adopen.Recordset.Fields(6)
sisaa3.Text = Val(sisaa.Text) - Val(sisaa2.Text)
outstok3.Text = Val(outstok.Text) + Val(outstok2.Text)
sisaa.Text = sisaa3.Text
outstok.Text = outstok3.Text
midoftgl.Text = Mid(tanggalaa.Text, 4, 2)
picket.Enabled = False
updatet.Enabled = True
ElseIf resep2.index.Text = "5" Then
sql = "select * from LPRSTOK where NMGenerik like '%" & resep2.nb5.Text & "%' order by NMGenerik asc"
adopen.RecordSource = sql
adopen.Refresh
idob.Text = adopen.Recordset.Fields(0)
namaa.Text = adopen.Recordset.Fields(1)
satuan.Text = adopen.Recordset.Fields(2)
jumlahaa.Text = adopen.Recordset.Fields(3)
tanggalaa.Text = adopen.Recordset.Fields(4)
sisaa.Text = adopen.Recordset.Fields(5)
sisaaa.Text = adopen.Recordset.Fields(5)
outstok.Text = adopen.Recordset.Fields(6)
sisaa3.Text = Val(sisaa.Text) - Val(sisaa2.Text)
outstok3.Text = Val(outstok.Text) + Val(outstok2.Text)
sisaa.Text = sisaa3.Text
outstok.Text = outstok3.Text
midoftgl.Text = Mid(tanggalaa.Text, 4, 2)
picket.Enabled = False
updatet.Enabled = True
Else
MsgBox "ERR0R !", vbCritical
End If
End Sub

Private Sub picket2_Timer()
sql = "select * from LPRSTOK where NMGenerik like '%" & inputstok.nm.Text & "%' order by NMGenerik asc"
adopen.RecordSource = sql
adopen.Refresh
idob.Text = adopen.Recordset.Fields(0)
namaa.Text = adopen.Recordset.Fields(1)
satuan.Text = adopen.Recordset.Fields(2)
jumlahaa.Text = adopen.Recordset.Fields(3)
tanggalaa.Text = adopen.Recordset.Fields(4)
sisaaa.Text = adopen.Recordset.Fields(5)
outstok.Text = adopen.Recordset.Fields(6)
picket2.Enabled = False
End Sub

Private Sub savet_Timer()
sql = "select * from LPRSTOK where IDObat ='" & idob.Text & "'"
 adopen.RecordSource = sql
 adopen.Refresh

 If adopen.Recordset.EOF Then
 adopen.Recordset.AddNew
 adopen.Recordset.Fields(0) = idob.Text
 adopen.Recordset.Fields(1) = namaa.Text
  adopen.Recordset.Fields(2) = satuan.Text
 adopen.Recordset.Fields(3) = jumlahaa.Text
 adopen.Recordset.Fields(4) = tanggalaa.Text
 adopen.Recordset.Fields(5) = sisaa.Text
 If stts.Caption = "ADA" Then
 adopen.Recordset.Fields(6) = outstok.Text
 ElseIf stts.Caption = "KOSONG" Then
 adopen.Recordset.Fields(6) = 0
 End If
 adopen.Recordset.Update
 adopen.Refresh
 deletet.Enabled = True
 savet.Enabled = False
 Else
deletet.Enabled = True
savet.Enabled = False
 End If
End Sub

Private Sub tgl1_Click()
If (Mid(noww.Text, 4, 1) > (Mid(now2.Text, 4, 1))) Then
MsgBox "SIlahkan Input data bulan baru"
Else
MsgBox "Errrr"
End If
End Sub
Private Sub Timer1_Timer()
noww.Text = Date
End Sub


Private Sub updatet_Timer()
If (Mid(noww.Text, 4, 2) > tanggalaa.Text) Then
MsgBox "SIlahkan Input data bulan baru"
Else
sql = "select * from LPRSTOK where IDObat ='" & idob.Text & "'"
adopen.RecordSource = sql
adopen.Recordset.Update
adopen.Recordset.Fields(1) = namaa.Text
adopen.Recordset.Fields(2) = satuan.Text
adopen.Recordset.Fields(3) = jumlahaa.Text
adopen.Recordset.Fields(4) = tanggalaa.Text
adopen.Recordset.Fields(5) = sisaa.Text
adopen.Recordset.Fields(6) = outstok.Text
adopen.Recordset.Update
adopen.Refresh
sql = " select * from LPRSTOK"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
updatet.Enabled = False
deletet.Enabled = True
End If
End Sub
