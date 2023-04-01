VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form riwayatstok 
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optket 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8820
      TabIndex        =   7
      Top             =   1440
      Width           =   1755
   End
   Begin VB.OptionButton opttgl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6540
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optgenerik 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NM Generik"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4020
      TabIndex        =   5
      Top             =   1440
      Width           =   1845
   End
   Begin VB.TextBox txtcari 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1230
      TabIndex        =   3
      Text            =   "Cari"
      Top             =   1920
      Width           =   10155
   End
   Begin Project1.jcbutton clox1 
      Height          =   405
      Left            =   14040
      TabIndex        =   1
      Top             =   90
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   714
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "X"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin MSAdodcLib.Adodc ador 
      Height          =   330
      Left            =   9120
      Top             =   6270
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Caption         =   " "
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
   Begin MSDataGridLib.DataGrid dgr 
      Height          =   3645
      Left            =   570
      TabIndex        =   2
      Top             =   2430
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   6429
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
      Caption         =   "R I W A Y A T"
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
   Begin Project1.jcbutton cetak 
      Height          =   555
      Left            =   12060
      TabIndex        =   4
      Top             =   1830
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   979
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   "&CETAK"
      ForeColor       =   16384
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -570
      Top             =   -30
      Width           =   855
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   1050
      Shape           =   4  'Rounded Rectangle
      Top             =   1770
      Width           =   10545
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   11790
      Shape           =   4  'Rounded Rectangle
      Top             =   1770
      Width           =   1665
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   -30
      Top             =   1290
      Width           =   14655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RIWAYAT OPERASI STOK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5910
      TabIndex        =   0
      Top             =   300
      Width           =   3945
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   0
      Top             =   1110
      Width           =   15255
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   -570
      Top             =   -210
      Width           =   15165
   End
End
Attribute VB_Name = "riwayatstok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cetak_Click()
Set drstk.DataSource = ador
drstk.Refresh
Load drstk
drstk.Show
End Sub

Private Sub clox1_Click()
Unload Me
menu.Show
End Sub

Private Sub Form_Load()
sql = "select * from RSTOK"
ador.RecordSource = sql
ador.Refresh
Set dgr.DataSource = ador
dgr.Refresh
dgr.Columns(5).Width = 3000
dgr.Columns(1).Width = 3000
End Sub
Private Sub Form_MouseDown(Button As Integer, _
Shift As Integer, x As Single, Y As Single)
Dim ReturnValue As Long
  If Button = 1 Then
     Call ReleaseCapture
     ReturnValue = SendMessage(Me.hwnd, _
     WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub


Private Sub txtcari_KeyPress(KeyAscii As Integer)
If optgenerik.Value Then
sql = "select * from rstok where NMGenerik like '%" & txtcari.Text & "%' order by NMGenerik asc"
ador.RecordSource = sql
ador.Refresh
ElseIf opttgl.Value Then
sql = "select * from rstok where tanggal like '%" & txtcari.Text & "%' order by nama_dokter asc"
ador.RecordSource = sql
ador.Refresh
ElseIf optket.Value Then
sql = "select * from rstok where keterangan_riwayat like '%" & txtcari.Text & "%' order by keterangan_riwayat asc"
ador.RecordSource = sql
ador.Refresh
Else
End If
End Sub
