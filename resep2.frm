VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form resep2 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12645
   Icon            =   "resep2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11970
      Top             =   690
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3915
      Left            =   960
      TabIndex        =   78
      Top             =   8220
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox waktu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7080
         TabIndex        =   91
         Top             =   1290
         Width           =   1695
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   9330
         Top             =   900
      End
      Begin VB.Timer Timer33 
         Interval        =   12
         Left            =   9720
         Top             =   900
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   9330
         Top             =   510
      End
      Begin VB.TextBox stok3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5340
         TabIndex        =   88
         Top             =   1290
         Width           =   1695
      End
      Begin VB.TextBox no_res 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   87
         Top             =   1290
         Width           =   1545
      End
      Begin VB.TextBox stok2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3630
         TabIndex        =   86
         Top             =   1290
         Width           =   1665
      End
      Begin VB.TextBox stok1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1950
         TabIndex        =   85
         Top             =   1290
         Width           =   1635
      End
      Begin VB.TextBox nmgenerik 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   82
         Top             =   750
         Width           =   2445
      End
      Begin VB.TextBox tgll 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5340
         TabIndex        =   81
         Top             =   750
         Width           =   2445
      End
      Begin VB.TextBox nmpasien 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2850
         TabIndex        =   80
         Top             =   750
         Width           =   2445
      End
      Begin VB.TextBox nomor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   8580
         TabIndex        =   79
         Top             =   150
         Width           =   1605
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   9720
         Top             =   510
      End
      Begin MSAdodcLib.Adodc ador 
         Height          =   330
         Left            =   6930
         Top             =   1800
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
         Height          =   1875
         Left            =   300
         TabIndex        =   83
         Top             =   1860
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   3307
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
         Caption         =   "riwayat"
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
      Begin VB.Label Label23 
         Caption         =   "Label15"
         Height          =   255
         Left            =   300
         TabIndex        =   90
         Top             =   1650
         Width           =   945
      End
      Begin VB.Label stts 
         BackStyle       =   0  'Transparent
         Caption         =   "Pemakaian ke 1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   420
         TabIndex        =   84
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   11580
      Top             =   5010
   End
   Begin VB.TextBox no5 
      Height          =   435
      Left            =   7260
      TabIndex        =   76
      Top             =   6450
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox no4 
      Height          =   435
      Left            =   7290
      TabIndex        =   75
      Top             =   5580
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox no3 
      Height          =   435
      Left            =   7320
      TabIndex        =   74
      Top             =   4710
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox no2 
      Height          =   435
      Left            =   7290
      TabIndex        =   73
      Top             =   3810
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox no1 
      Height          =   435
      Left            =   7260
      TabIndex        =   72
      Top             =   2970
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1215
      Left            =   3690
      TabIndex        =   57
      Top             =   7560
      Width           =   3285
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   65
         Top             =   270
         Width           =   195
      End
      Begin VB.Label nama2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   64
         Top             =   30
         Width           =   1305
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   63
         Top             =   30
         Width           =   195
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   62
         Top             =   570
         Width           =   195
      End
      Begin VB.Label umur2 
         BackStyle       =   0  'Transparent
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   61
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label cmbjk2 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   60
         Top             =   570
         Width           =   1305
      End
      Begin VB.Label alamat 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   59
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   58
         Top             =   840
         Width           =   195
      End
   End
   Begin VB.TextBox index 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   12330
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   8160
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5835
      Left            =   8430
      TabIndex        =   8
      Top             =   1830
      Width           =   2745
      Begin Project1.jcbutton adds 
         Height          =   885
         Left            =   0
         TabIndex        =   9
         Top             =   390
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1561
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
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "resep2.frx":0B9A
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton simpan 
         Height          =   915
         Left            =   0
         TabIndex        =   11
         Top             =   1260
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1614
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
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "resep2.frx":16A8
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton refreshh 
         Height          =   885
         Left            =   0
         TabIndex        =   14
         Top             =   3060
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1561
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "resep2.frx":1EC2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton CARI 
         Height          =   945
         Left            =   0
         TabIndex        =   15
         Top             =   3930
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1667
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Cari dan Cetak"
         ForeColor       =   32768
         PictureNormal   =   "resep2.frx":2AAC
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton hapus 
         Height          =   915
         Left            =   0
         TabIndex        =   17
         Top             =   2160
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1614
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "resep2.frx":3906
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton new 
         Height          =   975
         Left            =   0
         TabIndex        =   18
         Top             =   4860
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1720
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "resep2.frx":45A0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Navigasi"
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
         Left            =   840
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.Shape Shape17 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   585
         Left            =   0
         Top             =   0
         Width           =   3225
      End
   End
   Begin Project1.jcbutton x 
      Height          =   435
      Left            =   12240
      TabIndex        =   16
      Top             =   0
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5175
      Left            =   930
      TabIndex        =   26
      Top             =   2250
      Width           =   6045
      Begin VB.Frame f05 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   750
         TabIndex        =   66
         Top             =   4140
         Width           =   5175
         Begin VB.TextBox cara5 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   90
            TabIndex        =   70
            Top             =   540
            Width           =   5025
         End
         Begin VB.TextBox nb5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   150
            TabIndex        =   69
            Top             =   90
            Width           =   2865
         End
         Begin VB.TextBox j5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   3510
            TabIndex        =   68
            Top             =   90
            Width           =   465
         End
         Begin VB.TextBox j55 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   4050
            TabIndex        =   67
            Top             =   90
            Width           =   1065
         End
         Begin Project1.jcbutton clox5 
            Height          =   435
            Left            =   3060
            TabIndex        =   71
            Top             =   60
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   767
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
         Begin VB.Shape Shape14 
            Height          =   435
            Left            =   90
            Top             =   60
            Width           =   2955
         End
         Begin VB.Shape Shape13 
            Height          =   435
            Left            =   3480
            Top             =   60
            Width           =   1695
         End
      End
      Begin VB.Frame f04 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   750
         TabIndex        =   51
         Top             =   3270
         Width           =   5175
         Begin VB.TextBox j44 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   4050
            TabIndex        =   55
            Top             =   90
            Width           =   1065
         End
         Begin VB.TextBox j4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   3510
            TabIndex        =   54
            Top             =   90
            Width           =   465
         End
         Begin VB.TextBox nb4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   120
            TabIndex        =   53
            Top             =   90
            Width           =   2865
         End
         Begin VB.TextBox cara4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   90
            TabIndex        =   52
            Top             =   540
            Width           =   5025
         End
         Begin Project1.jcbutton clox4 
            Height          =   435
            Left            =   3060
            TabIndex        =   56
            Top             =   60
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   767
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
         Begin VB.Shape Shape12 
            Height          =   435
            Left            =   3480
            Top             =   60
            Width           =   1695
         End
         Begin VB.Shape Shape11 
            Height          =   435
            Left            =   90
            Top             =   60
            Width           =   2955
         End
      End
      Begin VB.Frame f03 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   750
         TabIndex        =   45
         Top             =   2400
         Width           =   5175
         Begin VB.TextBox cara3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   90
            TabIndex        =   49
            Top             =   540
            Width           =   5025
         End
         Begin VB.TextBox nb3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   120
            TabIndex        =   48
            Top             =   90
            Width           =   2865
         End
         Begin VB.TextBox j3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   3510
            TabIndex        =   47
            Top             =   90
            Width           =   465
         End
         Begin VB.TextBox j33 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   4050
            TabIndex        =   46
            Top             =   90
            Width           =   1065
         End
         Begin Project1.jcbutton clox3 
            Height          =   435
            Left            =   3060
            TabIndex        =   50
            Top             =   60
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   767
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
         Begin VB.Shape Shape10 
            Height          =   435
            Left            =   90
            Top             =   60
            Width           =   2955
         End
         Begin VB.Shape Shape9 
            Height          =   435
            Left            =   3480
            Top             =   60
            Width           =   1695
         End
      End
      Begin VB.Frame f02 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   750
         TabIndex        =   39
         Top             =   1530
         Width           =   5175
         Begin VB.TextBox j22 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   4050
            TabIndex        =   43
            Top             =   90
            Width           =   1065
         End
         Begin VB.TextBox j2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   3510
            TabIndex        =   42
            Top             =   90
            Width           =   465
         End
         Begin VB.TextBox nb2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   120
            TabIndex        =   41
            Top             =   90
            Width           =   2865
         End
         Begin VB.TextBox cara2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   90
            TabIndex        =   40
            Top             =   540
            Width           =   5025
         End
         Begin Project1.jcbutton clox2 
            Height          =   435
            Left            =   3060
            TabIndex        =   44
            Top             =   60
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   767
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
         Begin VB.Shape Shape8 
            Height          =   435
            Left            =   3480
            Top             =   60
            Width           =   1695
         End
         Begin VB.Shape Shape7 
            Height          =   435
            Left            =   90
            Top             =   60
            Width           =   2955
         End
      End
      Begin VB.Frame f01 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   750
         TabIndex        =   31
         Top             =   660
         Width           =   5175
         Begin VB.TextBox cara1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   90
            TabIndex        =   38
            Top             =   570
            Width           =   5055
         End
         Begin VB.TextBox nb1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   120
            TabIndex        =   34
            Top             =   90
            Width           =   2865
         End
         Begin VB.TextBox j1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   3510
            TabIndex        =   33
            Top             =   90
            Width           =   465
         End
         Begin VB.TextBox j11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   4050
            TabIndex        =   32
            Top             =   90
            Width           =   1065
         End
         Begin Project1.jcbutton clox1 
            Height          =   435
            Left            =   3060
            TabIndex        =   35
            Top             =   60
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   767
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
         Begin VB.Shape Shape5 
            Height          =   435
            Left            =   90
            Top             =   60
            Width           =   2955
         End
         Begin VB.Shape Shape6 
            Height          =   435
            Left            =   3480
            Top             =   60
            Width           =   1695
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2670
         TabIndex        =   37
         Top             =   -30
         Width           =   1305
      End
      Begin VB.Label tglresep2 
         BackStyle       =   0  'Transparent
         Caption         =   "tgl"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3690
         TabIndex        =   36
         Top             =   -60
         Width           =   1545
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   30
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Obat"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   900
         TabIndex        =   28
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   300
         Width           =   1155
      End
   End
   Begin MSDataGridLib.DataGrid dgpen 
      Height          =   1275
      Left            =   180
      TabIndex        =   13
      Top             =   10110
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   2249
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
      Left            =   10860
      Top             =   8460
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
   Begin VB.Image forb 
      Height          =   480
      Left            =   0
      Picture         =   "resep2.frx":56CA
      Stretch         =   -1  'True
      Top             =   60
      Width           =   525
   End
   Begin VB.Label proses 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   6000
      TabIndex        =   97
      Top             =   1470
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label bpjs 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor BPJS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3990
      TabIndex        =   96
      Top             =   7350
      Width           =   1305
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3690
      TabIndex        =   95
      Top             =   7320
      Width           =   195
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor BPJS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      TabIndex        =   94
      Top             =   7380
      Width           =   2445
   End
   Begin VB.Shape Shape18 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   9780
      Width           =   12705
   End
   Begin VB.Label tanggalutama 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   180
      TabIndex        =   93
      Top             =   9300
      Width           =   1665
   End
   Begin VB.Label waktuutama 
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   180
      TabIndex        =   92
      Top             =   9510
      Width           =   1665
   End
   Begin VB.Label Label20 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7260
      TabIndex        =   89
      Top             =   8610
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ER-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   870
      TabIndex        =   77
      Top             =   930
      Width           =   765
   End
   Begin VB.Shape Shape3 
      Height          =   7455
      Left            =   810
      Top             =   1440
      Width           =   6225
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "No Izin Praktek"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2190
      TabIndex        =   25
      Top             =   1860
      Width           =   2055
   End
   Begin VB.Label nodok 
      BackStyle       =   0  'Transparent
      Caption         =   "No Izin Praktek"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4650
      TabIndex        =   24
      Top             =   1830
      Width           =   2415
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4350
      TabIndex        =   23
      Top             =   1830
      Width           =   195
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Dokter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2190
      TabIndex        =   22
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label namadok 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Dokter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4650
      TabIndex        =   21
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4350
      TabIndex        =   20
      Top             =   1530
      Width           =   195
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      TabIndex        =   19
      Top             =   8460
      Width           =   2445
   End
   Begin VB.Image back 
      Height          =   600
      Left            =   30
      Picture         =   "resep2.frx":6519
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Shape shp1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   8190
      Top             =   1590
      Width           =   3225
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      TabIndex        =   7
      Top             =   8160
      Width           =   2445
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Umur"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   7860
      Width           =   2445
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pasien  "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      TabIndex        =   5
      Top             =   7590
      Width           =   2445
   End
   Begin VB.Label nomor2 
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   930
      Width           =   825
   End
   Begin VB.Shape Shape4 
      Height          =   945
      Left            =   810
      Top             =   510
      Width           =   1395
   End
   Begin VB.Label abu 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Resep :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   930
      TabIndex        =   3
      Top             =   570
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LUBUK PAKAM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3750
      TabIndex        =   2
      Top             =   1110
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JL. DIPONEGORO, PETAPAHAN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   840
      Width           =   3105
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PUSKESMAS LUBUK PAKAM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2940
      TabIndex        =   0
      Top             =   540
      Width           =   3105
   End
   Begin VB.Shape Shape2 
      Height          =   945
      Left            =   810
      Top             =   510
      Width           =   6225
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8595
      Left            =   690
      Top             =   390
      Width           =   6495
   End
   Begin VB.Shape Shape15 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "resep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub CARI_Click()
cresep.Show
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

Private Sub adds_Click()
f01.Visible = True
f02.Visible = True
f03.Visible = True
f04.Visible = True
f05.Visible = True
proses.Caption = "PROSES"
End Sub



Private Sub back_Click()
If proses.Caption = "PROSES" Then
forb.Visible = True
MsgBox "Anda Tidak Dapat Keluar Bila belum menyelesaikan resep", vbCritical, "W A R N I N G !!!"
Else
Unload Me
resep.Show
resep.adopen.RecordSource = "resep"
resep.adopen.Refresh
Set resep.dgpen.DataSource = resep.adopen
resep.dgpen.AllowUpdate = False
resep.dgpen.TabStop = False
resep.dgpen.Refresh
End If
End Sub



Private Sub clox1_Click()
f01.Visible = False
nb1.Text = ""
j1.Text = ""
End Sub

Private Sub clox2_Click()
f02.Visible = False
nb2.Text = ""
j2.Text = ""
End Sub

Private Sub clox3_Click()
f03.Visible = False
nb3.Text = ""
j3.Text = ""
End Sub

Private Sub clox4_Click()
f04.Visible = False
nb4.Text = ""
j4.Text = ""
End Sub

Private Sub clox5_Click()
f05.Visible = False
nb5.Text = ""
j5.Text = ""
End Sub

Private Sub Form_Load()
f01.Visible = False
f02.Visible = False
f03.Visible = False
f04.Visible = False
f05.Visible = False
sql = "select * from resep"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh


sql = "select * from rstok_item"
ador.RecordSource = sql
ador.Refresh
Set dgr.DataSource = ador
dgr.Refresh


Label20.Caption = adopen.Recordset.RecordCount
Label23.Caption = ador.Recordset.RecordCount
Timer5.Enabled = False
Timer33.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
waktuutama.Caption = Time
tanggalutama.Caption = Date
forb.Visible = False
End Sub

Private Sub hapus_Click()
If Label20.Caption = "0" Then
Timer5.Enabled = True
Else
adopen.Recordset.Delete
dgpen.Refresh
Label20.Caption = adopen.Recordset.RecordCount
End If
End Sub

Private Sub nb1_Click()
ribot.Show
index.Text = "1"
End Sub
Private Sub nb2_Click()
ribot.Show
index.Text = "2"
End Sub
Private Sub nb3_Click()
ribot.Show
index.Text = "3"
End Sub
Private Sub nb4_Click()
ribot.Show
index.Text = "4"
End Sub
Private Sub nb5_Click()
ribot.Show
index.Text = "5"
End Sub

Private Sub new_Click()
f01.Visible = False
f02.Visible = False
f03.Visible = False
f04.Visible = False
f05.Visible = False
nb1.Text = ""
j1.Text = ""
nb2.Text = ""
j2.Text = ""
nb3.Text = ""
j3.Text = ""
nb4.Text = ""
j4.Text = ""
nb5.Text = ""
j5.Text = ""
j11.Text = ""
j22.Text = ""
j33.Text = ""
j44.Text = ""
j55.Text = ""
cara1.Text = ""
cara2.Text = ""
cara3.Text = ""
cara4.Text = ""
cara5.Text = ""
End Sub

Private Sub refreshh_Click()
adopen.RecordSource = "resep"
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.AllowUpdate = False
dgpen.TabStop = False
dgpen.Refresh
Label20.Caption = adopen.Recordset.RecordCount
End Sub

Private Sub simpan_Click()
sql = "select * from resep where no_resep='" & nomor2.Caption & "'"
adopen.RecordSource = sql
adopen.Refresh
If adopen.Recordset.EOF Then
adopen.Recordset.AddNew
adopen.Recordset.Fields(0) = nomor2.Caption
adopen.Recordset.Fields(1) = tglresep2.Caption
adopen.Recordset.Fields(2) = nama2.Caption
adopen.Recordset.Fields(3) = umur2.Caption
adopen.Recordset.Fields(4) = cmbjk2.Caption
adopen.Recordset.Fields(5) = nb1.Text
adopen.Recordset.Fields(6) = j1.Text + j11.Text
adopen.Recordset.Fields(7) = nb2.Text
adopen.Recordset.Fields(8) = j2.Text + j22.Text
adopen.Recordset.Fields(9) = nb3.Text
adopen.Recordset.Fields(10) = j3.Text + j33.Text
adopen.Recordset.Fields(11) = nb4.Text
adopen.Recordset.Fields(12) = j4.Text + j44.Text
adopen.Recordset.Fields(13) = nb5.Text
adopen.Recordset.Fields(14) = j5.Text + j55.Text
adopen.Recordset.Fields(15) = namadok.Caption
adopen.Recordset.Fields(16) = nodok.Caption
adopen.Recordset.Fields(17) = cara1.Text
adopen.Recordset.Fields(18) = cara2.Text
adopen.Recordset.Fields(19) = cara3.Text
adopen.Recordset.Fields(20) = cara4.Text
adopen.Recordset.Fields(21) = cara5.Text
adopen.Recordset.Fields(22) = alamat.Caption
adopen.Recordset.Fields(23) = no1.Text
adopen.Recordset.Fields(24) = no2.Text
adopen.Recordset.Fields(25) = no3.Text
adopen.Recordset.Fields(26) = no4.Text
adopen.Recordset.Fields(27) = no5.Text
adopen.Recordset.Fields(28) = bpjs.Caption
adopen.Recordset.Update
adopen.Refresh
proses.Caption = "DONE"
adds.Enabled = False
forb.Visible = False
f01.Visible = False
f02.Visible = False
f03.Visible = False
f04.Visible = False
f05.Visible = False
nb1.Text = ""
j1.Text = ""
nb2.Text = ""
j2.Text = ""
nb3.Text = ""
j3.Text = ""
nb4.Text = ""
j4.Text = ""
nb5.Text = ""
j5.Text = ""
j11.Text = ""
j22.Text = ""
j33.Text = ""
j44.Text = ""
j55.Text = ""
cara1.Text = ""
cara2.Text = ""
cara3.Text = ""
cara4.Text = ""
cara5.Text = ""
MsgBox "Resep Telah disimpan,", vbInformation, "Information"
Else
MsgBox "Data Sudah Ada !", vbInformation, "Information"
MsgBox "Nomor Resep sudah ada sebelumnya, silahkan tambahkan ulang !", vbInformation, "Information"
End If
End Sub


Private Sub Timer1_Timer()
waktuutama.Caption = Time
tanggalutama.Caption = Date
End Sub

Private Sub Timer4_Timer()
nomor.Text = ""
nmgenerik.Text = ""
nmpasien.Text = ""
tgll.Text = ""
no_res.Text = ""
stok1.Text = ""
stok2.Text = ""
stok3.Text = ""
ador.RecordSource = "rstok_item"
ador.Refresh
Set dgr.DataSource = ador
dgr.AllowUpdate = False
dgr.TabStop = False
dgr.Refresh
Timer4.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
 If Label23.Caption = "0" Then
sql = "select * from rstok_item where no_riwayat='" & nomor.Text & "'"
ador.RecordSource = sql
ador.Refresh
If ador.Recordset.EOF Then

nomor.Text = "00001"

Timer2.Enabled = False
Else
MsgBox "Data Sudah Ada !", vbInformation, "Information"
MsgBox "Nomor Resep sudah ada sebelumnya, silahkan tambahkan ulang !", vbInformation, "Information"
End If


Else
sql = "select * from rstok_item where no_riwayat in(select max(no_riwayat) from rstok_item) order by no_riwayat"
ador.RecordSource = sql
ador.Refresh
Set dgr.DataSource = ador
dgr.Refresh
nomor.Text = ador.Recordset.Fields(0)
Timer3.Enabled = True
Timer2.Enabled = False
End If
End Sub


Private Sub Timer33_Timer()
sql = "select * from rstok_item where no_riwayat ='" & nomor.Text & "'"
ador.RecordSource = sql
ador.Refresh
If ador.Recordset.EOF Then
ador.Recordset.AddNew
ador.Recordset.Fields(0) = nomor.Text
ador.Recordset.Fields(1) = nmgenerik.Text
ador.Recordset.Fields(2) = nmpasien.Text
ador.Recordset.Fields(3) = tgll.Text
ador.Recordset.Fields(4) = no_res.Text
ador.Recordset.Fields(5) = stok1.Text
ador.Recordset.Fields(6) = stok2.Text
ador.Recordset.Fields(7) = stok3.Text
ador.Recordset.Fields(8) = waktu.Text
ador.Recordset.Update
ador.Refresh

Timer4.Enabled = True
Timer3.Enabled = True
Timer33.Enabled = False
Else
End If
End Sub

Private Sub Timer3_Timer()
Dim isi As Integer
Dim a
If (Mid(nomor.Text, 1, 1) <> 0) Then
isi = Right(nomor.Text, 5) + 1
Timer3.Enabled = False
If isi = 100000 Then
a = MsgBox("NILAI SUDAH MELEBIHI 19999!!! MERESET NILAI", vbInformation, "Informasi")
nomor.Text = "00001"
Timer3.Enabled = False
Else
nomor.Text = isi
Timer3.Enabled = False
End If
ElseIf (Mid(nomor.Text, 2, 1) <> 0) Then
isi = Right(nomor.Text, 4) + 1
Timer3.Enabled = False
If isi = 10000 Then
nomor.Text = "10000"
Timer3.Enabled = False
Else
nomor.Text = "0" & isi
Timer3.Enabled = False
End If
ElseIf (Mid(nomor.Text, 3, 1) <> 0) Then
isi = Right(nomor.Text, 3) + 1
Timer3.Enabled = False
If isi = 1000 Then
nomor.Text = "01000"
Timer3.Enabled = False
Else
nomor.Text = "00" & isi
Timer3.Enabled = False
End If
ElseIf (Mid(nomor.Text, 4, 1) <> 0) Then
isi = Right(nomor.Text, 2) + 1
Timer3.Enabled = False
If isi = 100 Then
nomor.Text = "00100"
Timer3.Enabled = False
Else
nomor.Text = "000" & isi
Timer3.Enabled = False
End If
ElseIf (Mid(nomor.Text, 5, 1) <> 0) Then
isi = Right(nomor.Text, 1) + 1
Timer3.Enabled = False
If isi = 10 Then
nomor.Text = "00010"
Timer3.Enabled = False
Else
nomor.Text = "0000" & isi
Timer3.Enabled = False
End If
End If
End Sub

Private Sub Timer5_Timer()
MsgBox "DATA TELAH KOSONG....", vbInformation, "P U S K E S M A S"
Label20.Caption = adopen.Recordset.RecordCount
Timer5.Enabled = False
End Sub


Private Sub x_Click()
Unload resep2
menu.Show
End Sub
