VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5475
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.Frame fus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1875
         Left            =   390
         TabIndex        =   7
         Top             =   2340
         Width           =   4305
         Begin VB.TextBox txtnama 
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
            Height          =   405
            Left            =   480
            TabIndex        =   8
            Top             =   1110
            Width           =   2985
         End
         Begin VB.Shape Shape20 
            BorderColor     =   &H00808080&
            Height          =   675
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   960
            Width           =   3885
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "U S E R N A M E"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   585
            Left            =   1050
            TabIndex        =   9
            Top             =   0
            Width           =   1995
         End
      End
      Begin VB.Frame fpw 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1755
         Left            =   390
         TabIndex        =   3
         Top             =   2310
         Width           =   4305
         Begin VB.TextBox txtpw 
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
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   450
            PasswordChar    =   "X"
            TabIndex        =   4
            Top             =   1170
            Width           =   2775
         End
         Begin VB.Shape Shape19 
            BorderColor     =   &H00808080&
            Height          =   675
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   990
            Width           =   3885
         End
         Begin VB.Image x 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3270
            Picture         =   "login.frx":0000
            Stretch         =   -1  'True
            Top             =   1110
            Width           =   465
         End
         Begin VB.Image o 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   3270
            Picture         =   "login.frx":0C28
            Stretch         =   -1  'True
            Top             =   1170
            Width           =   465
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "P A S S W O R D"
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
            Height          =   315
            Left            =   1020
            TabIndex        =   5
            Top             =   30
            Width           =   1995
         End
      End
      Begin Project1.jcbutton login 
         Height          =   465
         Left            =   1680
         TabIndex        =   2
         Top             =   4710
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
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
         BackColor       =   8454016
         Caption         =   "L O G I N"
         ForeColor       =   32768
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SILAHKAN LOGIN"
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
         Height          =   585
         Left            =   1410
         TabIndex        =   6
         Top             =   150
         Width           =   3135
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1950
         Picture         =   "login.frx":16ED
         Stretch         =   -1  'True
         Top             =   930
         Width           =   765
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   465
         Left            =   1740
         Picture         =   "login.frx":1F60
         Stretch         =   -1  'True
         Top             =   990
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4710
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4710
         Y1              =   4590
         Y2              =   4590
      End
   End
   Begin MSAdodcLib.Adodc adodata 
      Height          =   435
      Left            =   9060
      Top             =   5490
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
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
      Connect         =   $"login.frx":26F3
      OLEDBString     =   $"login.frx":2785
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
   Begin MSDataGridLib.DataGrid dgdata 
      Height          =   465
      Left            =   7500
      TabIndex        =   0
      Top             =   5490
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   820
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
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   1125
      Left            =   3750
      Top             =   5310
      Width           =   2625
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   1125
      Left            =   -60
      Top             =   5190
      Width           =   2625
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   300
      Width           =   4815
   End
   Begin VB.Image back 
      Height          =   450
      Left            =   2310
      Picture         =   "login.frx":2817
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   435
   End
   Begin VB.Shape shpp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   15
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1125
      Left            =   0
      Top             =   0
      Width           =   5025
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr, ctr2, R As Double
Dim ctr3 As String
Private Sub ext_Click()
Y = MsgBox("Apakah anda ingin keluar dari program?", vbYesNo + vbInformation, "Alert")
If Y = vbYes Then
Unload Me
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
End
End Sub

Private Sub Form_Load()
o.Visible = True
x.Visible = False
fpw.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Do
Me.Left = Me.Left + 40
Me.Move Me.Left, Me.Top
DoEvents
Loop Until Me.Left > Screen.Width
End Sub



Private Sub login_Click()
sql = "select * from login where username='" & txtnama.Text & "'"
sql = "select * from login where password='" & txtpw.Text & "'"
adodata.RecordSource = sql
adodata.Refresh
If adodata.Recordset.EOF Then
MsgBox "Username atau Password Yang Anda Masukkan Salah ! ", vbCritical, "ERROR 01"
fpw.Visible = False
fus.Visible = True
txtpw = ""
txtnama = ""
MsgBox "Daftar Terlebih Dahulu ! ", vbCritical, "ERROR 01"
ElseIf txtpw = "" Then
Timer3.Enabled = True
MsgBox "Username atau Password Yang Anda Masukkan Salah ! ", vbCritical, "ERROR 01"
fpw.Visible = False
fus.Visible = True
txtpw = ""
txtnama = ""
MsgBox "Masukkan Password Terlebih Dahulu ! ", vbCritical, "ERROR 01"
ElseIf txtnama.Text = "" Then
Timer2.Enabled = True
MsgBox "Username atau Password Yang Anda Masukkan Salah ! ", vbCritical, "ERROR 01"
fpw.Visible = False
fus.Visible = True
txtpw = ""
txtnama = ""
MsgBox "Masukkan Username Terlebih Dahulu ! ", vbCritical, "ERROR 01"
Else
menu.Show
Unload Me
End If
End Sub

Private Sub o_Click()
If txtpw.PasswordChar = "X" Then
txtpw.PasswordChar = Char
o.Visible = False
x.Visible = True
Else
MsgBox " error", vbYesNo + vbAbortRetryIgnore + vbCritical, "Beware OF THE DOG"
End If
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
fus.Visible = False
fpw.Visible = True
txtpw.SetFocus
End If
End Sub

Private Sub x_Click()
If txtpw.PasswordChar = Char Then
txtpw.PasswordChar = "X"
o.Visible = True
x.Visible = False
Else
MsgBox " error", vbYesNo + vbAbortRetryIgnore + vbCritical, "Beware OF THE DOG"
End If
End Sub
