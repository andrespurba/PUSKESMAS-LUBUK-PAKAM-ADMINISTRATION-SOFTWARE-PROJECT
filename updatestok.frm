VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form updatestok 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fback 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   405
      Left            =   570
      TabIndex        =   23
      Top             =   90
      Width           =   2175
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Kembali ke Menustok"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   24
         Top             =   90
         Width           =   1635
      End
      Begin VB.Shape Shape19 
         FillColor       =   &H8000000B&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   60
         Top             =   60
         Width           =   1785
      End
   End
   Begin VB.Frame fbaru 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4725
      Left            =   2880
      TabIndex        =   4
      Top             =   1050
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
         Left            =   2220
         TabIndex        =   11
         Top             =   3240
         Width           =   2415
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
         TabIndex        =   10
         Top             =   2610
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   1560
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker ed 
         Height          =   345
         Left            =   7680
         TabIndex        =   12
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
         Format          =   103022593
         CurrentDate     =   43737
      End
      Begin VB.Label notif 
         BackStyle       =   0  'Transparent
         Caption         =   "Silahkan Update Data"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4380
         TabIndex        =   29
         Top             =   4020
         Width           =   2085
      End
      Begin VB.Shape Shape18 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1380
         Shape           =   3  'Circle
         Top             =   420
         Width           =   135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Silahkan Perbaharui Data Dibawah ini sesuai dengan yang di butuhkan."
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
         Height          =   585
         Left            =   1560
         TabIndex        =   22
         Top             =   390
         Width           =   8175
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
         Left            =   330
         TabIndex        =   21
         Top             =   3240
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   2580
         Width           =   2295
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         Height          =   4545
         Left            =   150
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   10275
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   2730
         Width           =   2295
      End
      Begin VB.Label Label4 
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
         TabIndex        =   13
         Top             =   1530
         Width           =   2295
      End
      Begin VB.Shape Shape11 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   120
         Top             =   1530
         Width           =   2085
      End
      Begin VB.Shape Shape12 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   120
         Top             =   2130
         Width           =   2085
      End
      Begin VB.Shape Shape13 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   120
         Top             =   2730
         Width           =   2085
      End
      Begin VB.Shape Shape14 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   5610
         Top             =   930
         Width           =   2085
      End
      Begin VB.Shape Shape15 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   5610
         Top             =   1500
         Width           =   2085
      End
      Begin VB.Shape Shape16 
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
         Left            =   120
         Top             =   3240
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6825
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2445
      Begin Project1.jcbutton updates 
         Height          =   795
         Left            =   0
         TabIndex        =   25
         Top             =   2400
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1402
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         Caption         =   "U P D A T E"
         ForeColor       =   4210752
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton calls 
         Height          =   795
         Left            =   0
         TabIndex        =   26
         Top             =   3180
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1402
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         Caption         =   "C A L L  D A T A"
         ForeColor       =   16512
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton cleans 
         Height          =   795
         Left            =   0
         TabIndex        =   27
         Top             =   3960
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1402
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         Caption         =   "C L E A N"
         ForeColor       =   4210688
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Shape Shape20 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   0
         Top             =   2280
         Width           =   2625
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   -30
         Top             =   1890
         Width           =   2625
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Form Pembaruan Data Stok"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   585
         Left            =   240
         TabIndex        =   3
         Top             =   1380
         Width           =   2265
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   300
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   300
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "U P D A T E R"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   585
         Left            =   360
         TabIndex        =   2
         Top             =   810
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   -60
         Top             =   1890
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
         Left            =   600
         TabIndex        =   1
         Top             =   1920
         Width           =   1635
      End
   End
   Begin MSDataGridLib.DataGrid dgpenapotek 
      Height          =   1515
      Left            =   2850
      TabIndex        =   28
      Top             =   5880
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   2672
      _Version        =   393216
      BackColor       =   4210752
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
      Left            =   10230
      Top             =   7290
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
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   180
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   7350
      Shape           =   3  'Circle
      Top             =   180
      Width           =   255
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   6870
      Shape           =   3  'Circle
      Top             =   180
      Width           =   255
   End
   Begin VB.Image back 
      Height          =   600
      Left            =   60
      Picture         =   "updatestok.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   13965
   End
End
Attribute VB_Name = "updatestok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
menustok.Show
Unload Me
End Sub
Private Sub back_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
fback.Visible = True
End Sub

Private Sub calls_Click()
ribot4.Show
End Sub

Private Sub cleans_Click()
tgl.Text = ""
nama.Text = ""
harga = ""
ed.Value = 0
satuan.Text = ""
batch.Text = ""
idob.Text = ""
jumlah11.Text = ""
Label20.Caption = "Jumlah Obat"
pabrik.Text = ""
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
fback.Visible = False
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
fback.Visible = False
End Sub

Private Sub updates_Click()
sql = "select * from APOTEK where IDObat ='" & idob.Text & "'"
adopenapotek.RecordSource = sql
adopenapotek.Recordset.Update
adopenapotek.Recordset.Fields(0) = idob.Text
adopenapotek.Recordset.Fields(1) = nama.Text
adopenapotek.Recordset.Fields(2) = satuan.Text
adopenapotek.Recordset.Fields(3) = jumlah11.Text
adopenapotek.Recordset.Fields(4) = harga.Text
adopenapotek.Recordset.Fields(5) = ed.Value
adopenapotek.Recordset.Fields(6) = batch.Text
adopenapotek.Recordset.Fields(7) = pabrik.Text
adopenapotek.Refresh
sql = " select * from APOTEK"
adopenapotek.RecordSource = sql
adopenapotek.Refresh
Set dgpenapotek.DataSource = adopenapotek
dgpenapotek.Refresh
MsgBox "Data Anda Berhasil di update !", vbInformation, "PuskesLBK"
nama.Text = ""
harga = ""
ed.Value = 0
satuan.Text = ""
batch.Text = ""
idob.Text = ""
jumlah11.Text = ""
Label20.Caption = "Jumlah Obat"
pabrik.Text = ""
notif.Visible = False
End Sub


Private Sub Form_Load()
notif.Visible = False
sql = "select * from APOTEK"
adopenapotek.RecordSource = sql
adopenapotek.Refresh
Set dgpenapotek.DataSource = adopenapotek
dgpenapotek.Refresh
End Sub
