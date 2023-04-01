VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form cresep 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optpasien 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nama Pasien"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   960
      Width           =   1845
   End
   Begin VB.OptionButton optno 
      BackColor       =   &H00C0C0C0&
      Caption         =   "No.Resep"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton optdok 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nama Dokter"
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
      Left            =   7410
      TabIndex        =   5
      Top             =   960
      Width           =   1755
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
      Left            =   360
      TabIndex        =   0
      Text            =   "Cari"
      Top             =   1620
      Width           =   10155
   End
   Begin Project1.jcbutton clox1 
      Height          =   615
      Left            =   11610
      TabIndex        =   2
      Top             =   -30
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   1085
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
   Begin MSAdodcLib.Adodc adopen 
      Height          =   330
      Left            =   10800
      Top             =   5550
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
   Begin Project1.jcbutton cetak 
      Height          =   555
      Left            =   10770
      TabIndex        =   4
      Top             =   1530
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
   Begin MSDataGridLib.DataGrid dgpen 
      Height          =   3645
      Left            =   30
      TabIndex        =   1
      Top             =   2250
      Width           =   11985
      _ExtentX        =   21140
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
   Begin VB.Label s2 
      BackStyle       =   0  'Transparent
      Caption         =   "Silahkan tentukan item yang ingin di cari, sebelum mencetak !(OPSIONAL)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   675
      Left            =   3120
      TabIndex        =   8
      Top             =   690
      Width           =   7905
   End
   Begin VB.Shape s1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   405
      Left            =   2550
      Top             =   930
      Width           =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   4950
      X2              =   7290
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CETAK RESEP"
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
      Left            =   5310
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   10710
      Shape           =   4  'Rounded Rectangle
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   1470
      Width           =   10545
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   1170
      Shape           =   3  'Circle
      Top             =   180
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   720
      Shape           =   3  'Circle
      Top             =   180
      Width           =   255
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   240
      Shape           =   3  'Circle
      Top             =   180
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   0
      Top             =   1320
      Width           =   12075
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -60
      Top             =   0
      Width           =   12075
   End
End
Attribute VB_Name = "cresep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cetak_Click()
Set dresep.DataSource = adopen
dresep.Refresh
Load dresep
dresep.Show
End Sub

Private Sub clox1_Click()
resep2.Show
Unload cresep
End Sub

Private Sub Form_Load()
sql = "select * from RESEP"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
s1.Visible = True
txtcari.Enabled = False
txtcari.Text = "Tidak Dapat Mencari  X( .."
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

Private Sub optdok_Click()
s1.Visible = False
s2.Visible = False
txtcari.Text = "Cari"
txtcari.Enabled = True
End Sub

Private Sub optno_Click()
s1.Visible = False
s2.Visible = False
txtcari.Text = "Cari"
txtcari.Enabled = True
End Sub

Private Sub optpasien_Click()
s1.Visible = False
s2.Visible = False
txtcari.Text = "Cari"
txtcari.Enabled = True
End Sub

Private Sub txtcari_Click()
txtcari.Text = ""
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If optpasien.Value Then
sql = "select * from resep where nama_pasien like '%" & txtcari.Text & "%' order by nama_pasien asc"
adopen.RecordSource = sql
adopen.Refresh
ElseIf optdok.Value Then
sql = "select * from resep where nama_dokter like '%" & txtcari.Text & "%' order by nama_dokter asc"
adopen.RecordSource = sql
adopen.Refresh
ElseIf optno.Value Then
sql = "select * from resep where no_resep like '%" & txtcari.Text & "%' order by no_resep asc"
adopen.RecordSource = sql
adopen.Refresh
Else
End If
End Sub
