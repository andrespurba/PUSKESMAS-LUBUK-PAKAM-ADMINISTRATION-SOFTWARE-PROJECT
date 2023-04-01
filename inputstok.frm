VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form inputstok 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13140
   Icon            =   "inputstok.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   1260
      Top             =   7170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3015
      Left            =   0
      TabIndex        =   45
      Top             =   8760
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox jumlah222 
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
         Left            =   4020
         TabIndex        =   58
         Top             =   390
         Width           =   1755
      End
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
         Left            =   5820
         TabIndex        =   54
         Top             =   750
         Width           =   1755
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   8430
         Top             =   630
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   8850
         Top             =   570
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   9270
         Top             =   540
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
         Left            =   8940
         TabIndex        =   53
         Top             =   120
         Width           =   1245
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
         Left            =   2070
         TabIndex        =   49
         Top             =   750
         Width           =   1905
      End
      Begin VB.TextBox jumlah111 
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
         Left            =   4020
         TabIndex        =   48
         Top             =   750
         Width           =   1755
      End
      Begin VB.TextBox namaa 
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
         TabIndex        =   47
         Top             =   750
         Width           =   1665
      End
      Begin MSAdodcLib.Adodc ador 
         Height          =   330
         Left            =   6930
         Top             =   1800
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
         Left            =   330
         TabIndex        =   46
         Top             =   1140
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
      Begin VB.Label stts 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Baru"
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
         TabIndex        =   51
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   5625
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   2445
      Begin Project1.jcbutton neww 
         Height          =   675
         Left            =   60
         TabIndex        =   60
         Top             =   4050
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1191
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
         PictureNormal   =   "inputstok.frx":0AFE
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton deletee 
         Height          =   675
         Left            =   60
         TabIndex        =   61
         Top             =   3330
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1191
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
         PictureNormal   =   "inputstok.frx":160C
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton refreshh 
         Height          =   675
         Left            =   60
         TabIndex        =   62
         Top             =   4770
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1191
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
         PictureNormal   =   "inputstok.frx":19D2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton simpan 
         Height          =   675
         Left            =   60
         TabIndex        =   59
         Top             =   2610
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1191
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
         PictureNormal   =   "inputstok.frx":25BC
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton updatee 
         Height          =   675
         Left            =   60
         TabIndex        =   63
         Top             =   2610
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1191
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
         PictureNormal   =   "inputstok.frx":2DD6
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Image Image1 
         Height          =   1365
         Left            =   510
         Picture         =   "inputstok.frx":3ADC
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1515
      End
      Begin VB.Shape alert2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   735
         Left            =   30
         Top             =   2580
         Width           =   2385
      End
      Begin VB.Shape Shape6 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   -30
         Top             =   2490
         Width           =   2625
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Navigasi"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   630
         TabIndex        =   1
         Top             =   2010
         Width           =   1635
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   585
         Left            =   -60
         Top             =   1920
         Width           =   2625
      End
   End
   Begin Project1.jcbutton baru 
      Height          =   315
      Left            =   6210
      TabIndex        =   3
      Top             =   510
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "Jenis Baru"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton ada 
      Height          =   315
      Left            =   8040
      TabIndex        =   4
      Top             =   510
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Caption         =   "Yang Sudah Ada"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7605
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   10485
      Begin VB.Frame fbaru 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   4725
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   10485
         Begin VB.TextBox jumlah11 
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
            Left            =   7680
            TabIndex        =   30
            Top             =   3150
            Width           =   1905
         End
         Begin VB.TextBox idob 
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
            Left            =   7680
            TabIndex        =   29
            Top             =   2610
            Width           =   2445
         End
         Begin VB.TextBox tgl 
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
            Left            =   2190
            TabIndex        =   28
            Top             =   960
            Width           =   2445
         End
         Begin VB.TextBox pabrik 
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
            Left            =   7680
            TabIndex        =   27
            Top             =   1500
            Width           =   2445
         End
         Begin VB.TextBox batch 
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
            Left            =   7680
            TabIndex        =   26
            Top             =   930
            Width           =   2445
         End
         Begin VB.TextBox satuan 
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
            Left            =   2190
            TabIndex        =   25
            Top             =   2730
            Width           =   2445
         End
         Begin VB.TextBox harga 
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
            Left            =   2550
            TabIndex        =   24
            Top             =   2130
            Width           =   2085
         End
         Begin VB.TextBox nama 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   2220
            TabIndex        =   23
            Top             =   1560
            Width           =   2445
         End
         Begin MSComCtl2.DTPicker ed 
            Height          =   345
            Left            =   7680
            TabIndex        =   31
            Top             =   2040
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   32768
            CalendarTitleBackColor=   12632256
            CalendarTitleForeColor=   4210752
            CalendarTrailingForeColor=   8421631
            Format          =   2752513
            CurrentDate     =   43737
         End
         Begin Project1.jcbutton call 
            Height          =   225
            Left            =   6750
            TabIndex        =   32
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            ButtonStyle     =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   8454143
            Caption         =   "PANGGIL DATA"
            ForeColor       =   255
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Image k2 
            Height          =   225
            Left            =   9780
            Picture         =   "inputstok.frx":6246
            Stretch         =   -1  'True
            Top             =   3180
            Width           =   165
         End
         Begin VB.Image k1 
            Height          =   225
            Left            =   9780
            Picture         =   "inputstok.frx":78D4
            Stretch         =   -1  'True
            Top             =   3180
            Width           =   165
         End
         Begin VB.Label alert 
            BackStyle       =   0  'Transparent
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   465
            Left            =   10200
            TabIndex        =   50
            Top             =   3030
            Width           =   225
         End
         Begin VB.Shape Shape22 
            BackColor       =   &H008080FF&
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   9600
            Shape           =   4  'Rounded Rectangle
            Top             =   3120
            Width           =   525
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Obat"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5820
            TabIndex        =   44
            Top             =   3150
            Width           =   2295
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Rp"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   2220
            TabIndex        =   43
            Top             =   2100
            Width           =   2295
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "ID Obat"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5820
            TabIndex        =   42
            Top             =   2580
            Width           =   2295
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            Height          =   4545
            Left            =   150
            Shape           =   4  'Rounded Rectangle
            Top             =   90
            Width           =   10275
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Setelah Selesai melakukan inputting, gunakan tombol navigasi disamping ini. <------"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   675
            Left            =   300
            TabIndex        =   41
            Top             =   3750
            Width           =   7065
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Input Data Secara lengkap"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   3630
            TabIndex        =   40
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   330
            TabIndex        =   39
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Pabrik"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5820
            TabIndex        =   38
            Top             =   1500
            Width           =   2295
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5820
            TabIndex        =   37
            Top             =   930
            Width           =   2295
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Expire Date"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5880
            TabIndex        =   36
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   330
            TabIndex        =   35
            Top             =   2100
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   330
            TabIndex        =   34
            Top             =   2730
            Width           =   2295
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Generik"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   360
            TabIndex        =   33
            Top             =   1530
            Width           =   2295
         End
         Begin VB.Shape Shape5 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   120
            Top             =   960
            Width           =   2085
         End
         Begin VB.Shape Shape7 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   120
            Top             =   1530
            Width           =   2085
         End
         Begin VB.Shape Shape8 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   120
            Top             =   2130
            Width           =   2085
         End
         Begin VB.Shape Shape9 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   120
            Top             =   2730
            Width           =   2085
         End
         Begin VB.Shape Shape10 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   5610
            Top             =   930
            Width           =   2085
         End
         Begin VB.Shape Shape11 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   5610
            Top             =   1500
            Width           =   2085
         End
         Begin VB.Shape Shape12 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   5610
            Top             =   2040
            Width           =   2085
         End
         Begin VB.Shape Shape17 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   5610
            Top             =   2610
            Width           =   2085
         End
         Begin VB.Shape Shape21 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Left            =   5610
            Top             =   3150
            Width           =   2085
         End
      End
      Begin VB.Frame fada 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4065
         Left            =   30
         TabIndex        =   8
         Top             =   0
         Width           =   10305
         Begin VB.TextBox stok2 
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
            Height          =   420
            IMEMode         =   3  'DISABLE
            Left            =   8100
            TabIndex        =   13
            Top             =   1470
            Width           =   1725
         End
         Begin VB.TextBox tambah 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   4380
            TabIndex        =   12
            Text            =   "0"
            Top             =   2730
            Width           =   1155
         End
         Begin VB.TextBox stok1 
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
            Height          =   420
            IMEMode         =   3  'DISABLE
            Left            =   5910
            TabIndex        =   11
            Top             =   1470
            Width           =   1725
         End
         Begin VB.TextBox tgl2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   1320
            TabIndex        =   10
            Text            =   "tanggal"
            Top             =   3600
            Width           =   2445
         End
         Begin VB.TextBox nm 
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
            Height          =   420
            IMEMode         =   3  'DISABLE
            Left            =   570
            TabIndex        =   9
            Top             =   1470
            Width           =   3375
         End
         Begin Project1.jcbutton call2 
            Height          =   555
            Left            =   4260
            TabIndex        =   14
            Top             =   1410
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   979
            ButtonStyle     =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   8454143
            Caption         =   "PANGGIL DATA"
            ForeColor       =   255
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Image gembok2 
            Height          =   225
            Left            =   5820
            Picture         =   "inputstok.frx":8F17
            Stretch         =   -1  'True
            Top             =   2850
            Width           =   165
         End
         Begin VB.Image gembok1 
            Height          =   225
            Left            =   5820
            Picture         =   "inputstok.frx":A5A5
            Stretch         =   -1  'True
            Top             =   2850
            Width           =   165
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   1170
            TabIndex        =   21
            Top             =   3570
            Width           =   195
         End
         Begin VB.Label yakin 
            BackStyle       =   0  'Transparent
            Caption         =   "Jika Sudah yakin klik gembok hingga terkunci, lalu simpan!"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   465
            Left            =   6120
            TabIndex        =   20
            Top             =   2820
            Width           =   3645
         End
         Begin VB.Shape Shape20 
            BackColor       =   &H008080FF&
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   5730
            Shape           =   4  'Rounded Rectangle
            Top             =   2790
            Width           =   345
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Penambahan Stok"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   4080
            TabIndex        =   19
            Top             =   2370
            Width           =   1695
         End
         Begin VB.Shape Shape19 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   555
            Left            =   7860
            Top             =   1380
            Width           =   60
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Stok Terbaru"
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
            Left            =   8250
            TabIndex        =   18
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Shape Shape18 
            BackColor       =   &H008080FF&
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   645
            Left            =   4260
            Shape           =   4  'Rounded Rectangle
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Stok Sebelumnya"
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
            Left            =   5850
            TabIndex        =   17
            Top             =   1050
            Width           =   2655
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Generik"
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
            TabIndex        =   16
            Top             =   1020
            Width           =   2655
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   330
            TabIndex        =   15
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Shape sf 
            BackColor       =   &H008080FF&
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   585
            Left            =   7980
            Shape           =   4  'Rounded Rectangle
            Top             =   1380
            Width           =   1965
         End
         Begin VB.Shape Shape16 
            BackColor       =   &H008080FF&
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   585
            Left            =   5820
            Shape           =   4  'Rounded Rectangle
            Top             =   1380
            Width           =   1965
         End
         Begin VB.Shape Shape14 
            BackColor       =   &H008080FF&
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   585
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   1410
            Width           =   3825
         End
         Begin VB.Shape Shape13 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   3915
            Left            =   90
            Shape           =   4  'Rounded Rectangle
            Top             =   90
            Width           =   10155
         End
      End
      Begin VB.Frame fapotek 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2055
         Left            =   150
         TabIndex        =   6
         Top             =   4950
         Width           =   10485
         Begin MSDataGridLib.DataGrid dgpenapotek 
            Height          =   1815
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   10275
            _ExtentX        =   18124
            _ExtentY        =   3201
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
            Caption         =   "DATA STOK APOTEK"
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
         Begin MSAdodcLib.Adodc adopenapotek 
            Height          =   330
            Left            =   9360
            Top             =   2250
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
      End
   End
   Begin VB.Image back 
      Height          =   600
      Left            =   0
      Picture         =   "inputstok.frx":BBE8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label23 
      Caption         =   "Label23"
      Height          =   195
      Left            =   570
      TabIndex        =   57
      Top             =   6330
      Visible         =   0   'False
      Width           =   915
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
      Left            =   90
      TabIndex        =   55
      Top             =   7350
      Width           =   1665
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
      Left            =   90
      TabIndex        =   56
      Top             =   7140
      Width           =   1665
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
      Left            =   11490
      TabIndex        =   52
      Top             =   120
      Width           =   1635
   End
   Begin VB.Shape Shape15 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1485
      Left            =   0
      Top             =   6180
      Width           =   2475
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Penambahan Stok"
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
      Left            =   630
      TabIndex        =   2
      Top             =   120
      Width           =   3045
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   11070
      Top             =   90
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   13245
   End
End
Attribute VB_Name = "inputstok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Private Sub ada_Click()
ada.BackColor = &HFFFFFF
ada.ForeColor = &H8000&
baru.BackColor = &HC0C0C0
baru.ForeColor = &HFFFFFF

fada.Visible = True
fbaru.Visible = False

updatee.Visible = True
simpan.Visible = False

fapotek.Visible = True

stts.Caption = ada.Caption
selelaporan.Hide
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


Private Sub back_Click()
menustok.Show
Unload Me
menustok.adopen.RecordSource = "resep"
menustok.adopen.Refresh
Set menustok.dgpen.DataSource = menustok.adopen
menustok.dgpen.AllowUpdate = False
menustok.dgpen.TabStop = False
menustok.dgpen.Refresh
End Sub

Private Sub baru_Click()
ada.BackColor = &HC0C0C0
ada.ForeColor = &HFFFFFF
baru.BackColor = &HFFFFFF
baru.ForeColor = &H8000&

fada.Visible = False
fbaru.Visible = True

simpan.Visible = True
updatee.Visible = False
stts.Caption = baru.Caption
End Sub

Private Sub call_Click()
ribot2.Show
End Sub
Private Sub call2_Click()
ribot3.Show
End Sub

Private Sub deletee_Click()
MsgBox "Bila Ingin melakukan penghapusan data lakukan dari Menu (Stok) !", vbInformation, "PuskesLBK"
deletee.Enabled = False
End Sub


Private Sub Form_Load()
alert2.Visible = False
yakin.Visible = False
tgl.Text = Date
tgll.Text = Date
tgl2.Text = Date
gembok2.Visible = False
sql = "select * from APOTEK"
adopenapotek.RecordSource = sql
adopenapotek.Refresh
Set dgpenapotek.DataSource = adopenapotek
dgpenapotek.Refresh
sql = "select * from RSTOK"
ador.RecordSource = sql
ador.Refresh
Set dgr.DataSource = ador
dgr.Refresh
alert.Visible = False
stok2.Enabled = False
k2.Visible = False
updatee.Visible = False
tanggalutama.Caption = Date
waktuutama.Caption = Time
Label23.Caption = ador.Recordset.RecordCount
tgl.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
simpan.Enabled = False
tgl2.Enabled = False
End Sub





Private Sub gembok1_Click()
If gembok1.Visible = True Then
gembok1.Visible = False
gembok2.Visible = True
yakin.Visible = True
yakin.Caption = "Terkunci !"
tambah.Enabled = False
stok2.Text = Val(stok1.Text) + Val(tambah.Text)
tgll.Text = tgl2.Text
jumlah111.Text = stok1.Text
jumlah222.Text = stok2.Text
namaa.Text = nm.Text
waktu.Text = waktuutama.Caption
selelaporan.sisaa.Text = Val(tambah.Text) + Val(selelaporan.sisaaa)
Else
MsgBox "False Parse !", vbCritical
End If
End Sub

Private Sub gembok2_Click()
If gembok1.Visible = False Then
gembok1.Visible = True
gembok2.Visible = False
yakin.Visible = True
yakin.Caption = "Jika Sudah yakin klik gembok hingga terkunci, lalu UPDATE Data!"
tambah.Enabled = True
stok2.Text = stok1.Text
tambah.Text = "0"
tambah.SetFocus
jumlah111.Text = ""
namaa.Text = ""
waktu.Text = ""
Else
MsgBox "False Parse !", vbCritical
End If
End Sub




Private Sub idob_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
End Sub

Private Sub jumlah11_KeyPress(KeyAscii As Integer)
 If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
   alert.Visible = True
simpan.Enabled = False
End Sub

Private Sub k1_Click()
  If k1.Visible = True Then
k1.Visible = False
k2.Visible = True
alert.Visible = False
jumlah111.Text = jumlah11.Text
jumlah11.Enabled = False
simpan.Enabled = True
alert.Visible = True
waktu.Text = waktuutama.Caption
tgll.Text = Date
selelaporan.Hide
selelaporan.idob.Text = idob.Text
selelaporan.namaa.Text = nama.Text
selelaporan.jumlahaa.Text = jumlah11.Text
selelaporan.tanggalaa.Text = tgl.Text
selelaporan.sisaa.Text = jumlah11.Text
selelaporan.stts.Caption = "KOSONG"
selelaporan.satuan.Text = satuan.Text
Else
MsgBox "False Parse !", vbCritical
End If
End Sub

Private Sub k2_Click()
If k2.Visible = True Then
k2.Visible = False
k1.Visible = True
alert.Visible = False
jumlah111.Text = ""
jumlah11.Text = ""
jumlah11.Enabled = True
jumlah11.SetFocus
simpan.Enabled = False
Else
MsgBox "False Parse !", vbCritical
End If
End Sub

Private Sub neww_Click()
ribot2.Show
Timer2.Enabled = True
End Sub

Private Sub refreshh_Click()
adopenapotek.RecordSource = "APOTEK"
adopenapotek.Refresh
Set dgpenapotek.DataSource = adopenapotek
dgpenapotek.AllowUpdate = False
dgpenapotek.TabStop = False
dgpenapotek.Refresh
End Sub

Private Sub simpan_Click()
If nomor.Text = "00001" Then
sql = "select * from APOTEK where IDObat ='" & idob.Text & "'"
 adopenapotek.RecordSource = sql
 adopenapotek.Refresh

 If adopenapotek.Recordset.EOF Then
 adopenapotek.Recordset.AddNew
 adopenapotek.Recordset.Fields(0) = idob.Text
 adopenapotek.Recordset.Fields(1) = nama.Text
 adopenapotek.Recordset.Fields(2) = satuan.Text
 adopenapotek.Recordset.Fields(3) = jumlah11.Text
 adopenapotek.Recordset.Fields(4) = harga.Text
 adopenapotek.Recordset.Fields(5) = ed.Value
 adopenapotek.Recordset.Fields(6) = batch.Text
 adopenapotek.Recordset.Fields(7) = pabrik.Text
 adopenapotek.Recordset.Update
 adopenapotek.Refresh
 batch.Text = ""
 pabrik.Text = ""
 jumlah11.Text = ""

 k1.Visible = True
 k2.Visible = False
 jumlah11.Enabled = True
 selelaporan.savet.Enabled = True
 Else
 MsgBox "Item Telah Ditambahkan sebelumnya !", vbInformation, "INFO 001"
  k1.Visible = True
 k2.Visible = False
 End If

 sql = "select * from RSTOK where no_riwayat ='" & nomor.Text & "'"
 ador.RecordSource = sql
 ador.Refresh
 If ador.Recordset.EOF Then
 ador.Recordset.AddNew
 ador.Recordset.Fields(0) = nomor.Text
 ador.Recordset.Fields(1) = namaa.Text
 ador.Recordset.Fields(2) = jumlah111.Text
 ador.Recordset.Fields(3) = jumlah111.Text
 ador.Recordset.Fields(4) = tgll.Text
 ador.Recordset.Fields(5) = "Penambahan Stok Baru"
 ador.Recordset.Fields(6) = waktu.Text
 ador.Recordset.Update
 ador.Refresh
 Timer4.Enabled = True
Timer3.Enabled = True
  Unload Me
  inputstok.Show
 Else
 End If
Else
 sql = "select * from APOTEK where IDObat ='" & idob.Text & "'"
 adopenapotek.RecordSource = sql
 adopenapotek.Refresh

 If adopenapotek.Recordset.EOF Then
 adopenapotek.Recordset.AddNew
 adopenapotek.Recordset.Fields(0) = idob.Text
 adopenapotek.Recordset.Fields(1) = nama.Text
 adopenapotek.Recordset.Fields(2) = satuan.Text
 adopenapotek.Recordset.Fields(3) = jumlah11.Text
 adopenapotek.Recordset.Fields(4) = harga.Text
 adopenapotek.Recordset.Fields(5) = ed.Value
 adopenapotek.Recordset.Fields(6) = batch.Text
 adopenapotek.Recordset.Fields(7) = pabrik.Text
 adopenapotek.Recordset.Update
 adopenapotek.Refresh
 batch.Text = ""
 pabrik.Text = ""
 jumlah11.Text = ""

 k1.Visible = True
 k2.Visible = False
 jumlah11.Enabled = True
 selelaporan.savet.Enabled = True
 Else
 MsgBox "Item Telah Ditambahkan sebelumnya !", vbInformation, "INFO 001"
  k1.Visible = True
 k2.Visible = False
 End If

 sql = "select * from RSTOK where no_riwayat ='" & nomor.Text & "'"
 ador.RecordSource = sql
 ador.Refresh
 If ador.Recordset.EOF Then
 ador.Recordset.AddNew
 ador.Recordset.Fields(0) = nomor.Text
 ador.Recordset.Fields(1) = namaa.Text
 ador.Recordset.Fields(2) = jumlah111.Text
 ador.Recordset.Fields(3) = jumlah111.Text
 ador.Recordset.Fields(4) = tgll.Text
 ador.Recordset.Fields(5) = "Penambahan Stok Baru"
 ador.Recordset.Fields(6) = waktu.Text
 ador.Recordset.Update
 ador.Refresh
 Timer4.Enabled = True
Timer3.Enabled = True

 Else
  MsgBox "Item Telah Ditambahkan sebelumnya !", vbInformation, "INFO 001"
 End If
 End If

End Sub

Private Sub tambah_keypress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
yakin.Visible = True
End Sub



Private Sub tgl_click()
tgl.BorderStyle = 0
End Sub



Private Sub Timer2_Timer()
If Label23.Caption = "0" Then
sql = "select * from rstok where no_riwayat='" & nomor.Text & "'"
ador.RecordSource = sql
ador.Refresh
If ador.Recordset.EOF Then

nomor.Text = "00001"

Timer2.Enabled = False
Else
MsgBox "Data Sudah Ada !", vbInformation, "Information"
End If


Else
sql = "select * from rstok where no_riwayat in(select max(no_riwayat) from rstok) order by no_riwayat"
ador.RecordSource = sql
ador.Refresh
Set dgr.DataSource = ador
dgr.Refresh
nomor.Text = ador.Recordset.Fields(0)
Timer3.Enabled = True
Timer2.Enabled = False
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



Private Sub Timer4_Timer()
nomor.Text = ""
namaa.Text = ""
jumlah111.Text = ""
tgll.Text = ""
waktu.Text = ""
ador.RecordSource = "rstok"
ador.Refresh
Set dgr.DataSource = ador
dgr.AllowUpdate = False
dgr.TabStop = False
dgr.Refresh
Timer4.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer5_Timer()
tanggalutama.Caption = Date
waktuutama.Caption = Time
End Sub



Private Sub updatee_Click()
sql = "select * from APOTEK where IDObat ='" & idob.Text & "'"
adopenapotek.RecordSource = sql
adopenapotek.Recordset.Update
adopenapotek.Recordset.Fields(1) = nm.Text
adopenapotek.Recordset.Fields(3) = stok2.Text
adopenapotek.Recordset.Update
adopenapotek.Refresh
selelaporan.updatet.Enabled = True
gembok1.Visible = True
 gembok2.Visible = False
 yakin.Caption = "Jika Sudah yakin klik gembok hingga terkunci, lalu simpan!"
sql = " select * from APOTEK"
adopenapotek.RecordSource = sql
adopenapotek.Refresh
Set dgpenapotek.DataSource = adopenapotek
dgpenapotek.Refresh
nm.Text = ""
stok2.Text = ""
stok1.Text = ""
tambah.Text = ""


 sql = "select * from RSTOK where no_riwayat ='" & nomor.Text & "'"
 ador.RecordSource = sql
 ador.Refresh
 If ador.Recordset.EOF Then
 ador.Recordset.AddNew
 ador.Recordset.Fields(0) = nomor.Text
 ador.Recordset.Fields(1) = namaa.Text
 ador.Recordset.Fields(2) = jumlah111.Text
 ador.Recordset.Fields(3) = jumlah222.Text
 ador.Recordset.Fields(4) = tgll.Text
 ador.Recordset.Fields(5) = "Penambahan Stok Yang Telah Ada"
 ador.Recordset.Fields(6) = waktu.Text
 ador.Recordset.Update
 ador.Refresh
 Timer4.Enabled = True
Timer3.Enabled = True
 Else
 MsgBox "Item Telah Ditambahkan !", vbInformation, "INFO 001"
 End If
End Sub

