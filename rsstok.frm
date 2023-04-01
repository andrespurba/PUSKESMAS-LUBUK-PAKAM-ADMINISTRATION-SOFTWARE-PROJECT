VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rsstok 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   13860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   6435
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2445
      Begin Project1.jcbutton m1 
         Height          =   495
         Left            =   -30
         TabIndex        =   1
         Top             =   2310
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   873
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16384
         Caption         =   "Pengeluaran"
         ForeColor       =   8454016
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin Project1.jcbutton m2 
         Height          =   495
         Left            =   -30
         TabIndex        =   2
         Top             =   2820
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   873
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16384
         Caption         =   "Stok"
         ForeColor       =   8454016
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin Project1.jcbutton m3 
         Height          =   495
         Left            =   -30
         TabIndex        =   3
         Top             =   3330
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   873
         ButtonStyle     =   4
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16384
         Caption         =   "Laporan"
         ForeColor       =   8454016
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin Project1.jcbutton tambah 
         Height          =   495
         Left            =   -30
         TabIndex        =   4
         Top             =   3810
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   873
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16384
         Caption         =   "Kembali"
         ForeColor       =   8454016
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   -180
         Top             =   2250
         Width           =   2625
      End
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   510
         Picture         =   "rsstok.frx":0000
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1515
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   -60
         Top             =   1920
         Width           =   2625
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dashboard"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   630
         TabIndex        =   5
         Top             =   1950
         Width           =   1635
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   450
         Picture         =   "rsstok.frx":0BB7
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   375
      End
      Begin VB.Image ext 
         Height          =   315
         Left            =   1710
         Picture         =   "rsstok.frx":0FF9
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   345
      End
   End
   Begin MSDataGridLib.DataGrid dgpen 
      Height          =   3585
      Left            =   2790
      TabIndex        =   10
      Top             =   2490
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6324
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
      Left            =   13680
      Top             =   6570
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
   Begin Project1.jcbutton cmdprint 
      Height          =   585
      Left            =   3960
      TabIndex        =   13
      Top             =   6360
      Width           =   7365
      _ExtentX        =   3413
      _ExtentY        =   1032
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777088
      Caption         =   "Cetak Laporan"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Pemakaian Obat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   2790
      TabIndex        =   11
      Top             =   1050
      Width           =   6225
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   2430
      Top             =   1710
      Width           =   6165
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "APOTEK"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   12270
      TabIndex        =   8
      Top             =   120
      Width           =   1635
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   11730
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Laporan Bulanan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   3045
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   7020
      Width           =   13875
   End
   Begin VB.Label jam 
      BackStyle       =   0  'Transparent
      Caption         =   "Jam"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   12780
      TabIndex        =   7
      Top             =   7200
      Width           =   3045
   End
   Begin VB.Label tgl 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   12780
      TabIndex        =   6
      Top             =   7020
      Width           =   3045
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   13875
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Pemakaian Obat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   585
      Left            =   2820
      TabIndex        =   12
      Top             =   1080
      Width           =   6225
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   3810
      Shape           =   4  'Rounded Rectangle
      Top             =   6270
      Width           =   7755
   End
End
Attribute VB_Name = "rsstok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdprint_Click()
Set dtrlaporan.DataSource = adopen
dtrlaporan.Refresh
Load dtrlaporan
dtrlaporan.Show
End Sub

Private Sub Form_Load()
sql = "select * from LPRSTOK"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
End Sub

Private Sub m1_Click()
Unload Me
menu.Show
End Sub

Private Sub m2_Click()
Unload Me
menustok.Show
End Sub

Private Sub tambah_Click()
Unload Me
menu.Show
End Sub
