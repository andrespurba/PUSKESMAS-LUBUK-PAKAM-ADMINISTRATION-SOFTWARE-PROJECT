VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form resep 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14040
   FillColor       =   &H00FFFFFF&
   Icon            =   "resep.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   4530
      Top             =   840
   End
   Begin VB.OptionButton today 
      Caption         =   "Hari Ini"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5610
      TabIndex        =   37
      Top             =   3810
      Width           =   1095
   End
   Begin VB.OptionButton cust 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Custom Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6690
      TabIndex        =   36
      Top             =   3810
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4110
      Top             =   840
   End
   Begin VB.Frame fbayi 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   435
      Left            =   5640
      TabIndex        =   29
      Top             =   4980
      Width           =   2535
      Begin VB.TextBox angka 
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   30
         TabIndex        =   31
         Top             =   30
         Width           =   1155
      End
      Begin VB.ComboBox satuan 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   30
         Text            =   "Satuan"
         Top             =   30
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc adopen 
      Height          =   330
      Left            =   10380
      Top             =   360
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6765
      Left            =   2940
      TabIndex        =   2
      Top             =   1860
      Width           =   10755
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Data Obat"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   4155
         Left            =   5970
         TabIndex        =   15
         Top             =   960
         Width           =   4515
         Begin Project1.jcbutton jobatd 
            Height          =   3675
            Left            =   330
            TabIndex        =   16
            Top             =   270
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   6482
            ButtonStyle     =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            Caption         =   "Draf Resep"
            ForeColor       =   32768
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Data Pasien"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   5625
         Left            =   210
         TabIndex        =   4
         Top             =   1020
         Width           =   5145
         Begin VB.TextBox bpjs 
            BackColor       =   &H00C0C0C0&
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
            Left            =   2490
            TabIndex        =   43
            Top             =   5130
            Width           =   2445
         End
         Begin VB.Frame ftent 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   375
            Left            =   2430
            TabIndex        =   33
            Top             =   2100
            Width           =   2595
            Begin VB.OptionButton optdewasa 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Remaja - Dewasa"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   840
               TabIndex        =   35
               Top             =   0
               Width           =   1725
            End
            Begin VB.OptionButton optbayi 
               Caption         =   "Bayi "
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   60
               TabIndex        =   34
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.TextBox alamat 
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
            Left            =   2490
            TabIndex        =   24
            Top             =   3330
            Width           =   2445
         End
         Begin VB.TextBox nodok 
            BackColor       =   &H00C0C0C0&
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
            Left            =   2490
            TabIndex        =   21
            Top             =   4590
            Width           =   2445
         End
         Begin VB.ComboBox cmbdok 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   390
            Left            =   2490
            TabIndex        =   19
            Text            =   "Pilih --- >"
            Top             =   4020
            Width           =   2475
         End
         Begin VB.TextBox nomor 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   3600
            TabIndex        =   9
            Top             =   300
            Width           =   1305
         End
         Begin VB.TextBox tglresep 
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
            Left            =   2490
            TabIndex        =   8
            Top             =   960
            Width           =   2445
         End
         Begin VB.TextBox nama 
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
            Left            =   2460
            TabIndex        =   7
            Top             =   1560
            Width           =   2445
         End
         Begin VB.TextBox umur 
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
            Left            =   2460
            TabIndex        =   6
            Top             =   2130
            Width           =   2445
         End
         Begin VB.ComboBox cmbjk 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2460
            TabIndex        =   5
            Text            =   "Pilih --- >"
            Top             =   2790
            Width           =   2475
         End
         Begin Project1.jcbutton batal 
            Height          =   285
            Left            =   2490
            TabIndex        =   28
            Top             =   2490
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            ButtonStyle     =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   8421631
            Caption         =   "Cancel"
            ForeColor       =   128
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin Project1.jcbutton okay 
            Height          =   285
            Left            =   3660
            TabIndex        =   32
            Top             =   2490
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            ButtonStyle     =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   8454016
            Caption         =   "Confirm"
            ForeColor       =   32768
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin MSComCtl2.DTPicker dttgl 
            Height          =   345
            Left            =   2460
            TabIndex        =   38
            Top             =   930
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   103022593
            CurrentDate     =   43575
         End
         Begin Project1.jcbutton batal2 
            Height          =   255
            Left            =   2460
            TabIndex        =   39
            Top             =   1290
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            ButtonStyle     =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   8421631
            Caption         =   "Cancel"
            ForeColor       =   128
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor BPJS Pasien"
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
            Height          =   315
            Left            =   180
            TabIndex        =   42
            Top             =   5100
            Width           =   2295
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            Left            =   150
            TabIndex        =   25
            Top             =   3300
            Width           =   2295
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Izin Praktek"
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
            Height          =   315
            Left            =   180
            TabIndex        =   20
            Top             =   4560
            Width           =   2295
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Dokter"
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
            Height          =   315
            Left            =   180
            TabIndex        =   18
            Top             =   4020
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Resep"
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
            Left            =   180
            TabIndex        =   14
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Resep"
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
            Left            =   150
            TabIndex        =   13
            Top             =   930
            Width           =   2295
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pasien"
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
            Left            =   150
            TabIndex        =   12
            Top             =   1530
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Umur Pasien"
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
            Left            =   150
            TabIndex        =   11
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Kelamin"
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
            Left            =   150
            TabIndex        =   10
            Top             =   2790
            Width           =   2295
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0FFC0&
            FillStyle       =   0  'Solid
            Height          =   1815
            Left            =   -30
            Top             =   3840
            Width           =   5205
         End
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Data"
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
         Left            =   240
         TabIndex        =   17
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Silahkan Input Data Resep"
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
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   90
         Width           =   3015
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   90
         Width           =   255
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   9480
         Shape           =   3  'Circle
         Top             =   90
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   9930
         Shape           =   3  'Circle
         Top             =   90
         Width           =   255
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   465
         Left            =   -180
         Top             =   0
         Width           =   14055
      End
   End
   Begin VB.PictureBox adop 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11370
      ScaleHeight     =   435
      ScaleWidth      =   1785
      TabIndex        =   27
      Top             =   8820
      Visible         =   0   'False
      Width           =   1845
   End
   Begin MSDataGridLib.DataGrid dgdop 
      Height          =   1095
      Left            =   3390
      TabIndex        =   1
      Top             =   8850
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1931
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
   Begin MSDataGridLib.DataGrid dgpen 
      Height          =   885
      Left            =   5490
      TabIndex        =   26
      Top             =   690
      Visible         =   0   'False
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   1561
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
   Begin VB.Image back 
      Height          =   600
      Left            =   0
      Picture         =   "resep.frx":0B9A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label20 
      Caption         =   "Label15"
      Height          =   255
      Left            =   12720
      TabIndex        =   41
      Top             =   930
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   525
      Left            =   6420
      TabIndex        =   40
      Top             =   4170
      Width           =   1245
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Layanan Pasien"
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
      Height          =   495
      Left            =   570
      TabIndex        =   23
      Top             =   120
      Width           =   3315
   End
   Begin VB.Shape Shape13 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   1410
      Top             =   2130
      Width           =   330
   End
   Begin VB.Shape Shape12 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1530
      Top             =   2010
      Width           =   90
   End
   Begin VB.Label Label11 
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
      Left            =   12120
      TabIndex        =   22
      Top             =   120
      Width           =   1635
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   0
      Top             =   3060
      Width           =   2685
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   630
      Picture         =   "resep.frx":2644
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resep Obat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   3090
      TabIndex        =   0
      Top             =   1200
      Width           =   3315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   2940
      X2              =   13830
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   8145
      Left            =   -30
      Top             =   600
      Width           =   2715
   End
   Begin VB.Shape Shape11 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   11700
      Top             =   120
      Width           =   1875
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   0
      Top             =   -30
      Width           =   14055
   End
End
Attribute VB_Name = "resep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub angka_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
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
menu.Show
Unload Me
End Sub

Private Sub batal_Click()
ftent.Visible = True
fbayi.Visible = False
batal.Visible = False
umur.Text = ""
optbayi.Value = False
optdewasa.Value = False
batal.Visible = True
cmbjk.Enabled = False
alamat.Enabled = False
cmbdok.Enabled = False
nodok.Enabled = False
batal.Visible = False
okay.Visible = True
angka.Text = ""
End Sub





Private Sub batal2_Click()
batal2.Visible = False
cust.Visible = True
today.Visible = True
today.Value = False
cust.Value = False
End Sub

Private Sub cust_Click()
cust.Visible = False
today.Visible = False
batal2.Visible = True
tglresep.Visible = False
dttgl.Visible = True
End Sub



Private Sub nama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
umur.SetFocus
End If
End Sub

Private Sub okay_Click()
cmbjk.Enabled = True
alamat.Enabled = True
cmbdok.Enabled = True
nodok.Enabled = True

umur.Text = angka.Text + satuan.Text
ftent.Visible = False
okay.Visible = False
batal.Visible = True
fbayi.Visible = False
umur.Visible = True
End Sub

Private Sub optbayi_Click()
fbayi.Visible = True
ftent.Visible = False
angka.SetFocus
okay.Visible = True

cmbjk.Enabled = True
alamat.Enabled = False
cmbdok.Enabled = False
nodok.Enabled = False
satuan.Enabled = True
End Sub

Private Sub optdewasa_Click()
satuan.Text = "Tahun"
satuan.Enabled = False
fbayi.Visible = True
ftent.Visible = False
angka.SetFocus
okay.Visible = True

cmbjk.Enabled = False
alamat.Enabled = False
cmbdok.Enabled = False
nodok.Enabled = False
End Sub

Private Sub Form_Load()
ftent.Visible = True
fbayi.Visible = False
batal.Visible = False
umur.Visible = False
sql = "select * from resep"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
tglresep.Text = Date
cmbjk.AddItem "Laki-laki"
cmbjk.AddItem "Perempuan"
tglresep.Enabled = False
cmbdok.AddItem "dr. IMELDA PASARIBU"
cmbdok.AddItem "drg. JULIATI"
cmbdok.AddItem "dr. VERONIKA BRAHMANA"
cmbdok.AddItem "dr. DARI SIREGAR"
cmbdok.AddItem "dr. EVARINA"
cmbdok.AddItem "dr. JUANA LUSIANTI"
batal2.Visible = False
dttgl.Visible = False
satuan.AddItem " Hari"
satuan.AddItem " Bulan"
satuan.AddItem " Tahun"
nomor.Enabled = False
okay.Visible = False
Label20.Caption = adopen.Recordset.RecordCount
Timer3.Enabled = False
End Sub

Private Sub jobatd_Click()
resep2.Show
resep2.nodok.Caption = nodok.Text
resep2.namadok.Caption = cmbdok.Text
resep2.nama2.Caption = nama.Text
resep2.umur2.Caption = umur.Text
resep2.bpjs.Caption = bpjs.Text
If today.Value = True Then
resep2.tglresep2.Caption = tglresep.Text
ElseIf today.Value = False Then
resep2.tglresep2.Caption = dttgl.Value
End If
resep2.nomor2.Caption = nomor.Text
resep2.cmbjk2.Caption = cmbjk.Text
resep2.alamat.Caption = alamat.Text
Unload resep
resep2.adopen.RecordSource = "resep"
resep2.adopen.Refresh
Set resep2.dgpen.DataSource = resep2.adopen
resep2.dgpen.AllowUpdate = False
resep2.dgpen.TabStop = False
resep2.dgpen.Refresh
End Sub


Private Sub satuan_Click()
okay.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Label20.Caption = "0" Then
sql = "select * from resep where no_resep='" & nomor.Text & "'"
adopen.RecordSource = sql
adopen.Refresh
If adopen.Recordset.EOF Then

nomor.Text = "00001"
Timer1.Enabled = False
Else
MsgBox "Data Sudah Ada !", vbInformation, "Information"
MsgBox "Nomor Resep sudah ada sebelumnya, silahkan tambahkan ulang !", vbInformation, "Information"
End If

Else
sql = "select * from resep where no_resep in(select max(no_resep) from resep) order by no_resep"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
nomor.Text = adopen.Recordset.Fields(0)
Timer1.Enabled = False
Timer3.Enabled = True
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

Private Sub today_Click()
cust.Visible = False
today.Visible = False
batal2.Visible = True
tglresep.Visible = True
dttgl.Visible = False
End Sub
