VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ribot3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12030
   Icon            =   "ribot3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
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
      Left            =   330
      TabIndex        =   0
      Text            =   "Cari"
      Top             =   990
      Width           =   10575
   End
   Begin Project1.jcbutton cari 
      Height          =   585
      Left            =   11100
      TabIndex        =   1
      Top             =   870
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1032
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
      PictureNormal   =   "ribot3.frx":108A
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin MSDataGridLib.DataGrid dgpen 
      Height          =   3795
      Left            =   30
      TabIndex        =   2
      Top             =   1590
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   6694
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
   Begin Project1.jcbutton clox1 
      Height          =   405
      Left            =   11670
      TabIndex        =   3
      Top             =   0
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
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   11070
      Top             =   840
      Width           =   855
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   90
      Top             =   840
      Width           =   10935
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   0
      Top             =   690
      Width           =   12075
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
   Begin VB.Shape Shape8 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   10935
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   885
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA STOK APOTEK"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   90
      Width           =   3945
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   12075
   End
End
Attribute VB_Name = "ribot3"
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


Private Sub clox1_Click()
Unload Me
End Sub

Private Sub dgpen_Click()
inputstok.nm.Text = adopen.Recordset.Fields(1)
inputstok.stok1.Text = adopen.Recordset.Fields(3)
inputstok.stok2.Text = adopen.Recordset.Fields(3)
sql = "select * from APOTEK where NMGenerik like '%" & adopen.Recordset.Fields(1) & "%' order by NMGenerik asc"
inputstok.adopenapotek.RecordSource = sql
inputstok.adopenapotek.Refresh
inputstok.tambah.Enabled = True
inputstok.tambah.SetFocus
selelaporan.picket2.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
sql = "select * from APOTEK"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
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

Private Sub txtcari_Click()
txtcari.Text = ""
txtcari.SetFocus
End Sub

