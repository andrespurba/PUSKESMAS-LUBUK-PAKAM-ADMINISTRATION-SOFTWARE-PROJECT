VERSION 5.00
Begin VB.Form menu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13800
   Icon            =   "menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   13800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame jeniss 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   975
      Left            =   2430
      TabIndex        =   21
      Top             =   4410
      Width           =   2205
      Begin Project1.jcbutton use 
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
         Caption         =   "Pemakaian Stok"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin Project1.jcbutton operation 
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
         Caption         =   "Operasi Stok"
         ForeColor       =   4210752
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8820
      Top             =   7200
   End
   Begin VB.Frame fpo 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4725
      Left            =   2700
      TabIndex        =   6
      Top             =   1140
      Width           =   10785
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2595
         Left            =   7230
         TabIndex        =   14
         Top             =   900
         Width           =   2355
         Begin Project1.jcbutton P4 
            Height          =   1665
            Left            =   300
            TabIndex        =   15
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
            PictureNormal   =   "menu.frx":0D0A
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "PROGRAM PEMERINTAH"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   60
            TabIndex        =   16
            Top             =   2130
            Width           =   2265
         End
         Begin VB.Shape Shape12 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   300
            Top             =   2070
            Width           =   1785
         End
         Begin VB.Shape Shape13 
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2595
         Left            =   4290
         TabIndex        =   11
         Top             =   900
         Width           =   2355
         Begin Project1.jcbutton p3 
            Height          =   1665
            Left            =   300
            TabIndex        =   12
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
            PictureNormal   =   "menu.frx":1EC0
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Label Label7 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "POSKESDES"
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
            Left            =   570
            TabIndex        =   13
            Top             =   2100
            Width           =   1305
         End
         Begin VB.Shape Shape11 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   585
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   270
            Width           =   1455
         End
         Begin VB.Shape Shape10 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   300
            Top             =   2070
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2625
         Left            =   1350
         TabIndex        =   8
         Top             =   870
         Width           =   2325
         Begin Project1.jcbutton P1 
            Height          =   1665
            Left            =   300
            TabIndex        =   9
            Top             =   420
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
            Caption         =   ""
            PictureNormal   =   "menu.frx":2BDA
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin VB.Label Label3 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "PASIEN"
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
            Left            =   780
            TabIndex        =   10
            Top             =   2160
            Width           =   885
         End
         Begin VB.Shape Shape4 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   585
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   330
            Width           =   1455
         End
         Begin VB.Shape Shape5 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   300
            Top             =   2130
            Width           =   1785
         End
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   6210
         Top             =   570
         Width           =   315
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   4470
         Top             =   570
         Width           =   1635
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Pengeluaran Obat"
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
         Left            =   4590
         TabIndex        =   7
         Top             =   210
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   6435
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2445
      Begin Project1.jcbutton m1 
         Height          =   495
         Left            =   -30
         TabIndex        =   2
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
         TabIndex        =   4
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
         TabIndex        =   17
         Top             =   3330
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
         TabIndex        =   20
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
         Caption         =   "Riwayat"
         ForeColor       =   8454016
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Image ext 
         Height          =   315
         Left            =   1710
         Picture         =   "menu.frx":3784
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   345
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   450
         Picture         =   "menu.frx":3E9B
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   375
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
         TabIndex        =   3
         Top             =   1950
         Width           =   1635
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
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   510
         Picture         =   "menu.frx":42DD
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1515
      End
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
      Left            =   12750
      TabIndex        =   19
      Top             =   7020
      Width           =   3045
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
      Left            =   12750
      TabIndex        =   18
      Top             =   7200
      Width           =   3045
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -30
      Top             =   7020
      Width           =   13875
   End
   Begin VB.Shape sts 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2430
      Top             =   3930
      Width           =   105
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
      Left            =   12240
      TabIndex        =   5
      Top             =   120
      Width           =   1635
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   11790
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Utama"
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
      Left            =   210
      TabIndex        =   0
      Top             =   120
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
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_MouseDown(Button As Integer, _
Shift As Integer, x As Single, Y As Single)
Dim ReturnValue As Long
  If Button = 1 Then
     Call ReleaseCapture
     ReturnValue = SendMessage(Me.hwnd, _
     WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub



Private Sub ext_Click()
End
End Sub

Private Sub Form_Load()
sts.Visible = False
tgl.Caption = Date
jam.Caption = Time
jeniss.Visible = False
End Sub



Private Sub m1_Click()
sts.Top = 2910
sts.Visible = True
End Sub

Private Sub m2_Click()
sts.Top = 3450
sts.Visible = True
menustok.Show
menu.Hide

End Sub

Private Sub m3_Click()
sts.Top = 3930
sts.Visible = True
rsstok.Show
End Sub

Private Sub p2_Click()
resep.Show
menu.Hide
End Sub

Private Sub operation_Click()
riwayatstok.Show
Unload Me
End Sub

Private Sub P1_Click()
resep.Show
Unload Me
End Sub

Private Sub p3_Click()
MsgBox "Menunggu Dukungan Dari Pemerintah :)", vbInformation, "ANIMO Z"
End Sub

Private Sub P4_Click()
MsgBox "Menunggu Dukungan Dari Pemerintah :)", vbInformation, "ANIMO Z"
End Sub

Private Sub riwayat_Click()
riwayatstok.Show
Unload Me
End Sub

Private Sub tambah_Click()
If tambah.Caption = "Riwayat" Then
tambah.Caption = "-------->"
jeniss.Visible = True
ElseIf tambah.Caption = "-------->" Then
jeniss.Visible = False
tambah.Caption = "Riwayat"
Else
End If
End Sub

Private Sub Timer1_Timer()
tgl.Caption = Date
jam.Caption = Time
End Sub

Private Sub use_Click()
riwayatresep.Show
Unload Me
End Sub
