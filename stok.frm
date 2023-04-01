VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form menustok 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13875
   Icon            =   "stok.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   2790
      Top             =   810
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3015
      Left            =   2760
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   10335
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
         TabIndex        =   34
         Top             =   750
         Width           =   1665
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
         TabIndex        =   33
         Top             =   750
         Width           =   1755
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
         TabIndex        =   32
         Top             =   750
         Width           =   1905
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
         TabIndex        =   31
         Top             =   120
         Width           =   1245
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   9270
         Top             =   540
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   8850
         Top             =   570
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   8430
         Top             =   630
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
         TabIndex        =   30
         Top             =   750
         Width           =   1755
      End
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
         TabIndex        =   29
         Top             =   390
         Width           =   1755
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
         TabIndex        =   35
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
         TabIndex        =   36
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame jeniss 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   975
      Left            =   2430
      TabIndex        =   21
      Top             =   3000
      Width           =   2205
      Begin Project1.jcbutton ada 
         Height          =   495
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
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
         Caption         =   "Yang Telah Ada"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton belum 
         Height          =   495
         Left            =   0
         TabIndex        =   23
         Top             =   480
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
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
         BackColor       =   8454016
         Caption         =   "Jenis baru"
         ForeColor       =   4210752
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
   End
   Begin VB.Frame tampilstok 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5205
      Left            =   2670
      TabIndex        =   16
      Top             =   1290
      Width           =   11055
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
         Left            =   240
         TabIndex        =   18
         Text            =   "Cari"
         Top             =   570
         Width           =   8385
      End
      Begin Project1.jcbutton cari 
         Height          =   555
         Left            =   10140
         TabIndex        =   19
         Top             =   480
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   979
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "stok.frx":0D0A
         CaptionEffects  =   0
      End
      Begin MSAdodcLib.Adodc adopen 
         Height          =   330
         Left            =   9690
         Top             =   5850
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
      Begin MSDataGridLib.DataGrid dgpen 
         Height          =   4005
         Left            =   60
         TabIndex        =   17
         Top             =   1140
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   7064
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
      Begin Project1.jcbutton hapuss 
         Height          =   645
         Left            =   10140
         TabIndex        =   24
         Top             =   1200
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   979
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "stok.frx":1B64
         CaptionEffects  =   0
      End
      Begin Project1.jcbutton fress 
         Height          =   645
         Left            =   10140
         TabIndex        =   25
         Top             =   2100
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   979
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   ""
         PictureNormal   =   "stok.frx":27FE
         CaptionEffects  =   0
      End
      Begin Project1.jcbutton updatee 
         Height          =   675
         Left            =   10110
         TabIndex        =   26
         Top             =   3030
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         ButtonStyle     =   4
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
         PictureNormal   =   "stok.frx":33E8
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H008080FF&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   10050
         Shape           =   4  'Rounded Rectangle
         Top             =   2910
         Width           =   975
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H008080FF&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   10050
         Shape           =   4  'Rounded Rectangle
         Top             =   1980
         Width           =   945
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H008080FF&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   765
         Left            =   10050
         Shape           =   4  'Rounded Rectangle
         Top             =   1140
         Width           =   945
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   6120
         Top             =   360
         Width           =   315
      End
      Begin VB.Shape Shape6 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   4380
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sisa Stok"
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
         Height          =   375
         Left            =   4890
         TabIndex        =   20
         Top             =   0
         Width           =   2085
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H008080FF&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   10050
         Shape           =   4  'Rounded Rectangle
         Top             =   420
         Width           =   945
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H008080FF&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   420
         Width           =   9975
      End
   End
   Begin VB.Frame about 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4665
      Left            =   6180
      TabIndex        =   14
      Top             =   1740
      Width           =   3945
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "tambah,tampilkan dan cetak laporan stok obat."
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
         Height          =   735
         Left            =   0
         TabIndex        =   15
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Image i3 
         Height          =   3075
         Left            =   540
         Picture         =   "stok.frx":40EE
         Stretch         =   -1  'True
         Top             =   90
         Width           =   3195
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   6825
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2445
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1350
         Top             =   5850
      End
      Begin Project1.jcbutton tambah 
         Height          =   495
         Left            =   -30
         TabIndex        =   1
         Top             =   2400
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   873
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Tambah Stok"
         ForeColor       =   32768
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin Project1.jcbutton tampil 
         Height          =   525
         Left            =   -30
         TabIndex        =   2
         Top             =   3000
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   926
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Tampilkan Stok"
         ForeColor       =   32768
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "Label23"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   5730
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label21 
         Caption         =   "Label21"
         Height          =   375
         Left            =   210
         TabIndex        =   27
         Top             =   5310
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   600
         TabIndex        =   3
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   -60
         Top             =   1890
         Width           =   2625
      End
      Begin VB.Image i1 
         Height          =   1365
         Left            =   570
         Picture         =   "stok.frx":6858
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame fis 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   5625
      Left            =   3180
      TabIndex        =   4
      Top             =   1050
      Width           =   9885
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2595
         Left            =   1980
         TabIndex        =   8
         Top             =   1410
         Width           =   2355
         Begin Project1.jcbutton gudang 
            Height          =   1665
            Left            =   300
            TabIndex        =   9
            Top             =   360
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   2937
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
            Caption         =   " "
            PictureNormal   =   "stok.frx":8FC2
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Label Label11 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "GUDANG"
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
            Height          =   315
            Left            =   690
            TabIndex        =   10
            Top             =   2100
            Width           =   1305
         End
         Begin VB.Shape Shape19 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   300
            Top             =   2070
            Width           =   1785
         End
         Begin VB.Shape Shape20 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   585
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   270
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2595
         Left            =   5310
         TabIndex        =   5
         Top             =   1410
         Width           =   2355
         Begin Project1.jcbutton puskes 
            Height          =   1665
            Left            =   330
            TabIndex        =   6
            Top             =   360
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   2937
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
            Caption         =   " "
            PictureNormal   =   "stok.frx":9CDC
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Label Label10 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "PUSKESMAS"
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
            Height          =   315
            Left            =   540
            TabIndex        =   7
            Top             =   2100
            Width           =   1305
         End
         Begin VB.Shape Shape17 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   300
            Top             =   2070
            Width           =   1785
         End
         Begin VB.Shape Shape18 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   585
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   270
            Width           =   1455
         End
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Stok Obat"
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
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   180
         Width           =   2085
      End
      Begin VB.Shape Shape21 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   3720
         Top             =   540
         Width           =   1635
      End
      Begin VB.Shape Shape16 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   5460
         Top             =   540
         Width           =   315
      End
   End
   Begin VB.Image back 
      Height          =   600
      Left            =   0
      Picture         =   "stok.frx":A9F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ANIMO Z"
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
      Left            =   12240
      TabIndex        =   12
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok Obat"
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
      Left            =   540
      TabIndex        =   13
      Top             =   120
      Width           =   3045
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   11820
      Top             =   120
      Width           =   1875
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
End
Attribute VB_Name = "menustok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dgpen_Click()
namaa.Text = adopen.Recordset.Fields(1)
jumlah111.Text = adopen.Recordset.Fields(3)
jumlah222.Text = "KOSONG"
stts.Caption = "Penghapusan Data"
waktu.Text = Label6.Caption
hapuss.Enabled = True
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



Private Sub ada_Click()
selelaporan.Hide
gudang.Enabled = True
puskes.Enabled = True
jeniss.Visible = False
tambah.Caption = "Tambah Stok"

inputstok.ada.BackColor = &HFFFFFF
inputstok.ada.ForeColor = &H8000&
inputstok.baru.BackColor = &HC0C0C0
inputstok.baru.ForeColor = &HFFFFFF

inputstok.fada.Visible = True
inputstok.fbaru.Visible = False

inputstok.updatee.Visible = True
inputstok.simpan.Visible = False

inputstok.fapotek.Visible = True

inputstok.adopenapotek.RecordSource = "APOTEK"
inputstok.adopenapotek.Refresh
Set inputstok.dgpenapotek.DataSource = inputstok.adopenapotek
inputstok.dgpenapotek.AllowUpdate = False
inputstok.dgpenapotek.TabStop = False
inputstok.dgpenapotek.Refresh
End Sub

Private Sub back_Click()
Unload Me
menu.Show
End Sub

Private Sub belum_Click()
gudang.Enabled = True
puskes.Enabled = True
jeniss.Visible = False
tambah.Caption = "Tambah Stok"

inputstok.ada.BackColor = &HC0C0C0
inputstok.ada.ForeColor = &HFFFFFF
inputstok.baru.BackColor = &HFFFFFF
inputstok.baru.ForeColor = &H8000&

inputstok.fada.Visible = False
inputstok.fbaru.Visible = True

inputstok.simpan.Visible = True
inputstok.updatee.Visible = False


inputstok.adopenapotek.RecordSource = "APOTEK"
inputstok.adopenapotek.Refresh
Set inputstok.dgpenapotek.DataSource = inputstok.adopenapotek
inputstok.dgpenapotek.AllowUpdate = False
inputstok.dgpenapotek.TabStop = False
inputstok.dgpenapotek.Refresh
End Sub

Private Sub hapuss_Click()
If Label21.Caption = "0" Then
Timer1.Enabled = True
Else
adopen.Recordset.Delete
dgpen.Refresh '
selelaporan.adopen.Recordset.Delete
selelaporan.dgpen.Refresh
Label21.Caption = adopen.Recordset.RecordCount
selelaporan.Label21.Caption = selelaporan.adopen.Recordset.RecordCount
If nomor.Text = "00001" Then
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
 ador.Recordset.Fields(5) = "Penghapusan Data"
 ador.Recordset.Fields(6) = waktu.Text
 ador.Recordset.Update
 ador.Refresh
 Timer4.Enabled = True
 Timer3.Enabled = True
  Unload Me
  menustok.Show
 Else
 End If
Else
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
 ador.Recordset.Fields(5) = "Penghapusan Data"
 ador.Recordset.Fields(6) = waktu.Text
 ador.Recordset.Update
 ador.Refresh
 Timer4.Enabled = True
 Timer3.Enabled = True

 Else
 MsgBox "Item Telah Ditambahkan sebelumnya !", vbInformation, "INFO 001"
 End If
 End If
End If
End Sub

Private Sub fress_Click()
adopen.RecordSource = "APOTEK"
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.AllowUpdate = False
dgpen.TabStop = False
dgpen.Refresh
Label21.Caption = adopen.Recordset.RecordCount
End Sub

Private Sub puskes_Click()
inputstok.Show
Unload Me
End Sub
Private Sub Timer1_Timer()
MsgBox "DATA TELAH KOSONG....", vbInformation, "P U S K E S M A S"
Timer1.Enabled = False
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

Private Sub Timer5_Timer()
Label6.Caption = Time
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If txtcari.Text = "Cari" Then
txtcari.Text = ""
Else
sql = "select * from APOTEK where NMGenerik like '%" & txtcari.Text & "%' order by NMGenerik asc"
adopen.RecordSource = sql
adopen.Refresh
End If
End Sub

Private Sub Form_Load()
Label6.Caption = Timer
Timer1.Enabled = False
jeniss.Visible = False
i1.Visible = False
fis.Visible = False
tampilstok.Visible = False
sql = "select * from APOTEK"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
gudang.Enabled = False
puskes.Enabled = False
Label21.Caption = adopen.Recordset.RecordCount
hapuss.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
sql = "select * from RSTOK"
ador.RecordSource = sql
ador.Refresh
Set dgr.DataSource = ador
dgr.Refresh
Label23.Caption = ador.Recordset.RecordCount
tgll.Text = Date
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtcari.Text = "Cari"
adopen.RecordSource = "APOTEK"
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.AllowUpdate = False
dgpen.TabStop = False
dgpen.Refresh
Label21.Caption = adopen.Recordset.RecordCount
hapuss.Enabled = False
End Sub
Private Sub gudang_Click()
MsgBox "Menunggu Dukungan Dari Pemerintah :)", vbInformation, "ANIMO Z"
End Sub

Private Sub tambah_Click()
If tambah.Caption = "Tambah Stok" Then
tampilstok.Visible = False
tambah.Caption = "-------->"
jeniss.Visible = True
i1.Visible = True
fis.Visible = True
about.Visible = False
gudang.Enabled = False
puskes.Enabled = False
ElseIf tambah.Caption = "-------->" Then
jeniss.Visible = False
tambah.Caption = "Tambah Stok"
tampilstok.Visible = False
Else
End If
End Sub

Private Sub tampil_Click()
adopen.RecordSource = "APOTEK"
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.AllowUpdate = False
dgpen.TabStop = False
dgpen.Refresh
tampilstok.Visible = True
fis.Visible = False
jeniss.Visible = False
tambah.Caption = "Tambah Stok"
selelaporan.Hide
End Sub

Private Sub updatee_Click()
updatestok.Show
Unload Me
End Sub
