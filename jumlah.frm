VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form jumlah 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   3315
   ClientTop       =   2685
   ClientWidth     =   5055
   Icon            =   "jumlah.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox namakomp 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   480
      TabIndex        =   9
      Top             =   810
      Width           =   3045
   End
   Begin VB.TextBox jumlahkomp 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      TabIndex        =   8
      Top             =   810
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4050
      Top             =   2370
   End
   Begin VB.TextBox angka2 
      Height          =   465
      Left            =   4050
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2850
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox angka 
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
      Left            =   690
      TabIndex        =   2
      Top             =   1650
      Width           =   3585
   End
   Begin Project1.jcbutton clox1 
      Height          =   435
      Left            =   4650
      TabIndex        =   0
      Top             =   -30
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
   Begin Project1.jcbutton kurang 
      Height          =   435
      Left            =   1740
      TabIndex        =   4
      Top             =   2130
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   767
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
      BackColor       =   16777152
      Caption         =   "KONFIRMASI"
      ForeColor       =   4210688
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Project1.jcbutton batal 
      Height          =   435
      Left            =   1740
      TabIndex        =   5
      Top             =   2700
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   767
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
      BackColor       =   12632319
      Caption         =   "BATAL"
      ForeColor       =   128
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin MSAdodcLib.Adodc adopen 
      Height          =   330
      Left            =   12600
      Top             =   3300
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
   Begin MSDataGridLib.DataGrid dgpen 
      Height          =   2925
      Left            =   5610
      TabIndex        =   6
      Top             =   300
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   5159
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok"
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
      Left            =   3810
      TabIndex        =   11
      Top             =   480
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NMGenerik"
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
      TabIndex        =   10
      Top             =   480
      Width           =   2115
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   -30
      Top             =   1200
      Width           =   5085
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   0
      Top             =   3690
      Width           =   5085
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   780
      Top             =   1560
      Width           =   3435
   End
   Begin VB.Label angka0 
      BackStyle       =   0  'Transparent
      Caption         =   "Angka"
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
      Left            =   2130
      TabIndex        =   3
      Top             =   1230
      Width           =   2115
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
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
      Left            =   2070
      TabIndex        =   1
      Top             =   30
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   -30
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "jumlah"
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



Function Convert(IntegerAngka As Integer) As String
Dim I%
Dim IntegerSepuluh%, IntegerLima%, IntegerSatu%
Dim StringSeribu$, StringLimaRatus$
Dim StringSeratus$, StringLimaPuluh$
Dim StringSepuluh$, StringLima$, StringSatu$
Dim StringRom$

IntegerSatu = IntegerAngka
IntegerSepuluh = IntegerSatu \ 10
IntegerSatu = IntegerAngka Mod 10
IntegerLima = IntegerSatu \ 5
IntegerSatu = IntegerAngka Mod 5

For I = 0 To IntegerSepuluh - 1
StringSepuluh = StringSepuluh + "X"
Next

If IntegerSatu <> 4 Then
For I = 0 To IntegerLima - 1
StringLima = StringLima + "V"

Next
End If

For I = 0 To IntegerSatu - 1
StringSatu = StringSatu + "I"
Next

If IntegerSatu = 4 Then
If IntegerLima = 1 Then
StringSatu = StringRom + "IX"
Else
StringSatu = StringRom + "IV"
End If
End If

StringRom = StringSepuluh + StringLima + StringSatu
Convert = StringRom
End Function

Private Sub batal_Click()
If resep2.index.Text = "1" Then
resep2.nb1.Text = ""
resep2.j1.Text = ""
resep2.j11.Text = ""
Unload Me
ElseIf resep2.index.Text = "2" Then
resep2.nb2.Text = ""
resep2.j2.Text = ""
resep2.j22.Text = ""
Unload Me
ElseIf resep2.index.Text = "3" Then
resep2.nb3.Text = ""
resep2.j3.Text = ""
resep2.j33.Text = ""
Unload Me
ElseIf resep2.index.Text = "4" Then
resep2.nb4.Text = ""
resep2.j4.Text = ""
resep2.j44.Text = ""
Unload Me
ElseIf resep2.index.Text = "5" Then
resep2.nb5.Text = ""
resep2.j5.Text = ""
resep2.j55.Text = ""
Unload Me
Else
MsgBox "FALSE PARSE !)"
End If
End Sub

Private Sub clox1_Click()
Unload Me
ribot.Show
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
sql = "select * from APOTEK"
adopen.RecordSource = sql
adopen.Refresh
Set dgpen.DataSource = adopen
dgpen.Refresh
End Sub


Private Sub kurang_Click()
hitung = Val(jumlahkomp.Text) - Val(angka.Text)
If hitung < 0 Then
MsgBox "Data Stok Item yang anda ingin kan tidak mencukupi", vbYesNo + vbInformation, "APOTEK"
angka.Text = ""
angka.SetFocus
Else
If resep2.index.Text = "1" Then
 x = MsgBox("Apakah kamu sudah yakin ?", vbYesNo + vbInformation, "APOTEK")
 angka2.Text = Val(adopen.Recordset.Fields(3)) - Val(angka.Text)
 resep2.j1.Text = Convert(angka.Text)
 resep2.no1.Text = "NO"
 
 resep2.nmgenerik.Text = resep2.nb1.Text
 resep2.nmpasien.Text = resep2.nama2.Caption
 resep2.tgll.Text = resep2.tglresep2.Caption
 resep2.no_res.Text = resep2.nomor2.Caption
 resep2.stok1.Text = jumlahkomp.Text
 resep2.stok3.Text = angka2.Text
 resep2.stok2.Text = angka.Text
 resep2.waktu.Text = resep2.waktuutama
 resep2.clox1.Visible = False
 
 
 If x = vbYes Then
 If angka2.Text < 0 Then
 adopen.Recordset.Fields(3) = 0
 Timer1.Enabled = False
 MsgBox "Data Stok Item yang anda ingin kan telah kosong", vbYesNo + vbInformation, "APOTEK"
 x = MsgBox("Silahkan Input Data Stok Item", vbYesNo + vbInformation, "APOTEK")
  If x = vbYes Then
  jumlah.Hide
  resep2.Hide
  menustok.Show
  Else
  Unload Me
  End If
 Else
  resep2.Timer33.Enabled = True
 Timer1.Enabled = True
 End If
 ElseIf x = vbNo Then
 resep2.j1.Text = ""
 Timer1.Enabled = False
 angka.Text = ""
 
 Else
 MsgBox "FALSE PARSE !)", "APOTEK"
 End If
ElseIf resep2.index.Text = "2" Then
  x = MsgBox("Apakah kamu sudah yakin ?", vbYesNo + vbInformation, "APOTEK")
 angka2.Text = Val(adopen.Recordset.Fields(3)) - Val(angka.Text)
  resep2.j2.Text = Convert(angka.Text)
 resep2.no2.Text = "NO"
 
 resep2.nmgenerik.Text = resep2.nb2.Text
 resep2.nmpasien.Text = resep2.nama2.Caption
 resep2.tgll.Text = resep2.tglresep2.Caption
 resep2.no_res.Text = resep2.nomor2.Caption
 resep2.stok1.Text = jumlahkomp.Text
 resep2.stok3.Text = angka2.Text
 resep2.stok2.Text = angka.Text
 resep2.waktu.Text = resep2.waktuutama
 resep2.clox2.Visible = False
  
 If x = vbYes Then
 If angka2.Text < 0 Then
 adopen.Recordset.Fields(3) = 0
 Timer1.Enabled = False
 MsgBox "Data Stok Item yang anda ingin kan telah kosong", vbYesNo + vbInformation, "APOTEK"
 x = MsgBox("Silahkan Input Data Stok Item", vbYesNo + vbInformation, "APOTEK")
  If x = vbYes Then
  jumlah.Hide
  resep2.Hide
  menustok.Show
  Else
  Unload Me
  End If
 Else
  resep2.Timer33.Enabled = True
 Timer1.Enabled = True
 End If
 ElseIf x = vbNo Then
 resep2.j2.Text = ""
 Timer1.Enabled = False
 angka.Text = ""
 
 Else
 MsgBox "FALSE PARSE !)", "APOTEK"
 End If
ElseIf resep2.index.Text = "3" Then
 x = MsgBox("Apakah kamu sudah yakin ?", vbYesNo + vbInformation, "APOTEK")
 angka2.Text = Val(adopen.Recordset.Fields(3)) - Val(angka.Text)
  resep2.j3.Text = Convert(angka.Text)
 resep2.no3.Text = "NO"
 
 resep2.nmgenerik.Text = resep2.nb3.Text
 resep2.nmpasien.Text = resep2.nama2.Caption
 resep2.tgll.Text = resep2.tglresep2.Caption
 resep2.no_res.Text = resep2.nomor2.Caption
 resep2.stok1.Text = jumlahkomp.Text
 resep2.stok3.Text = angka2.Text
 resep2.stok2.Text = angka.Text
 resep2.waktu.Text = resep2.waktuutama
 resep2.clox3.Visible = False
 
 If x = vbYes Then
 If angka2.Text < 0 Then
 adopen.Recordset.Fields(3) = 0
 Timer1.Enabled = False
 MsgBox "Data Stok Item yang anda ingin kan telah kosong", vbYesNo + vbInformation, "APOTEK"
 x = MsgBox("Silahkan Input Data Stok Item", vbYesNo + vbInformation, "APOTEK")
  If x = vbYes Then
  jumlah.Hide
  resep2.Hide
  menustok.Show
  Else
  Unload Me
  End If
 Else
  resep2.Timer33.Enabled = True
 Timer1.Enabled = True
 End If
 ElseIf x = vbNo Then
 resep2.j3.Text = ""
 Timer1.Enabled = False
 angka.Text = ""
 
 Else
 MsgBox "FALSE PARSE !)", "APOTEK"
 End If
ElseIf resep2.index.Text = "4" Then
  x = MsgBox("Apakah kamu sudah yakin ?", vbYesNo + vbInformation, "APOTEK")
 angka2.Text = Val(adopen.Recordset.Fields(3)) - Val(angka.Text)
  resep2.j4.Text = Convert(angka.Text)
 resep2.no4.Text = "NO"
 resep2.nmgenerik.Text = resep2.nb4.Text
 resep2.nmpasien.Text = resep2.nama2.Caption
 resep2.tgll.Text = resep2.tglresep2.Caption
 resep2.no_res.Text = resep2.nomor2.Caption
 resep2.stok1.Text = jumlahkomp.Text
 resep2.stok3.Text = angka2.Text
 resep2.stok2.Text = angka.Text
 resep2.waktu.Text = resep2.waktuutama
 resep2.clox4.Visible = False
  
 If x = vbYes Then
 If angka2.Text < 0 Then
 adopen.Recordset.Fields(3) = 0
 Timer1.Enabled = False
 MsgBox "Data Stok Item yang anda ingin kan telah kosong", vbYesNo + vbInformation, "APOTEK"
 x = MsgBox("Silahkan Input Data Stok Item", vbYesNo + vbInformation, "APOTEK")
  If x = vbYes Then
  jumlah.Hide
  resep2.Hide
  menustok.Show
  Else
  Unload Me
  End If
 Else
  resep2.Timer33.Enabled = True
 Timer1.Enabled = True
 End If
 ElseIf x = vbNo Then
 resep2.j4.Text = ""
 Timer1.Enabled = False
 angka.Text = ""
 
 Else
 MsgBox "FALSE PARSE !)", "APOTEK"
 End If
ElseIf resep2.index.Text = "5" Then
  x = MsgBox("Apakah kamu sudah yakin ?", vbYesNo + vbInformation, "APOTEK")
 angka2.Text = Val(adopen.Recordset.Fields(3)) - Val(angka.Text)
  resep2.j5.Text = Convert(angka.Text)
 resep2.no5.Text = "NO"
 resep2.nmgenerik.Text = resep2.nb5.Text
 resep2.nmpasien.Text = resep2.nama2.Caption
 resep2.tgll.Text = resep2.tglresep2.Caption
 resep2.no_res.Text = resep2.nomor2.Caption
 resep2.stok1.Text = jumlahkomp.Text
 resep2.stok3.Text = angka2.Text
 resep2.stok2.Text = angka.Text
 resep2.waktu.Text = resep2.waktuutama
 resep2.clox5.Visible = False
 
 If x = vbYes Then
 If angka2.Text < 0 Then
 adopen.Recordset.Fields(3) = 0
 Timer1.Enabled = False
 MsgBox "Data Stok Item yang anda ingin kan telah kosong", vbYesNo + vbInformation, "APOTEK"
 x = MsgBox("Silahkan Input Data Stok Item", vbYesNo + vbInformation, "APOTEK")
  If x = vbYes Then
  jumlah.Hide
  resep2.Hide
  menustok.Show
  Else
  Unload Me
  End If
 Else
  resep2.Timer33.Enabled = True
 Timer1.Enabled = True
 End If
 ElseIf x = vbNo Then
 resep2.j5.Text = ""
 Timer1.Enabled = False
 angka.Text = ""
 
 Else
 MsgBox "FALSE PARSE !)", "APOTEK"
 End If


End If
End If
End Sub

Private Sub Text1_Change()

End Sub



Private Sub Timer1_Timer()
If resep2.index.Text = "1" Then
 If angka0.Caption = "Angka" Then
 angka0.Caption = "Tunggu."


 ElseIf angka0.Caption = "Tunggu." Then
 angka0.Caption = "Tunggu.."
 sql = "select * from APOTEK where NMGenerik like '%" & resep2.nb1.Text & "%' order by NMGenerik asc"
 adopen.RecordSource = sql
 adopen.Recordset.Update
 adopen.Recordset.Fields(3) = angka2.Text
 adopen.Recordset.Update
 adopen.Refresh
 sql = " select * from APOTEK"
 adopen.RecordSource = sql
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.Refresh
 resep2.x.Enabled = False
 selelaporan.Hide
 selelaporan.picket.Enabled = True
 selelaporan.sisaa2.Text = angka.Text
 selelaporan.outstok2.Text = angka.Text
 
 ElseIf angka0.Caption = "Tunggu.." Then
 angka0.Caption = "Tunggu..."

 ElseIf angka0.Caption = "Tunggu..." Then
 angka0.Caption = "Selesai"
 adopen.RecordSource = "APOTEK"
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.AllowUpdate = False
 dgpen.TabStop = False
 dgpen.Refresh

 Unload Me
 Else
 MsgBox "FALSE PARSE !)"
 End If
ElseIf resep2.index.Text = "2" Then
 If angka0.Caption = "Angka" Then
 angka0.Caption = "Tunggu."


 ElseIf angka0.Caption = "Tunggu." Then
 angka0.Caption = "Tunggu.."
 sql = "select * from APOTEK where NMGenerik like '%" & resep2.nb2.Text & "%' order by NMGenerik asc"
 adopen.RecordSource = sql
 adopen.Recordset.Update
 adopen.Recordset.Fields(3) = angka2.Text
 adopen.Recordset.Update
 adopen.Refresh
 sql = " select * from APOTEK"
 adopen.RecordSource = sql
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.Refresh
  resep2.x.Enabled = False
selelaporan.Hide
 selelaporan.picket.Enabled = True
 selelaporan.sisaa2.Text = angka.Text
 selelaporan.outstok2.Text = angka.Text
 ElseIf angka0.Caption = "Tunggu.." Then
 angka0.Caption = "Tunggu..."

 ElseIf angka0.Caption = "Tunggu..." Then
 angka0.Caption = "Selesai"
 adopen.RecordSource = "APOTEK"
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.AllowUpdate = False
 dgpen.TabStop = False
 dgpen.Refresh

 Unload Me
 Else
 MsgBox "FALSE PARSE !)"
 End If
ElseIf resep2.index.Text = "3" Then
 If angka0.Caption = "Angka" Then
 angka0.Caption = "Tunggu."


 ElseIf angka0.Caption = "Tunggu." Then
 angka0.Caption = "Tunggu.."
 sql = "select * from APOTEK where NMGenerik like '%" & resep2.nb3.Text & "%' order by NMGenerik asc"
 adopen.RecordSource = sql
 adopen.Recordset.Update
 adopen.Recordset.Fields(3) = angka2.Text
 adopen.Recordset.Update
 adopen.Refresh
 sql = " select * from APOTEK"
 adopen.RecordSource = sql
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.Refresh
  resep2.x.Enabled = False
selelaporan.Hide
 selelaporan.picket.Enabled = True
 selelaporan.sisaa2.Text = angka.Text
 selelaporan.outstok2.Text = angka.Text
 ElseIf angka0.Caption = "Tunggu.." Then
 angka0.Caption = "Tunggu..."

 ElseIf angka0.Caption = "Tunggu..." Then
 angka0.Caption = "Selesai"
 adopen.RecordSource = "APOTEK"
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.AllowUpdate = False
 dgpen.TabStop = False
 dgpen.Refresh

 Unload Me
 Else
 MsgBox "FALSE PARSE !)"
 End If
ElseIf resep2.index.Text = "4" Then
 If angka0.Caption = "Angka" Then
 angka0.Caption = "Tunggu."


 ElseIf angka0.Caption = "Tunggu." Then
 angka0.Caption = "Tunggu.."
 sql = "select * from APOTEK where NMGenerik like '%" & resep2.nb4.Text & "%' order by NMGenerik asc"
 adopen.RecordSource = sql
 adopen.Recordset.Update
 adopen.Recordset.Fields(3) = angka2.Text
 adopen.Recordset.Update
 adopen.Refresh
 sql = " select * from APOTEK"
 adopen.RecordSource = sql
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.Refresh
  resep2.x.Enabled = False
selelaporan.Hide
 selelaporan.picket.Enabled = True
 selelaporan.sisaa2.Text = angka.Text
 selelaporan.outstok2.Text = angka.Text
 ElseIf angka0.Caption = "Tunggu.." Then
 angka0.Caption = "Tunggu..."

 ElseIf angka0.Caption = "Tunggu..." Then
 angka0.Caption = "Selesai"
 adopen.RecordSource = "APOTEK"
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.AllowUpdate = False
 dgpen.TabStop = False
 dgpen.Refresh

 Unload Me
 Else
 MsgBox "FALSE PARSE !)"
 End If
ElseIf resep2.index.Text = "5" Then
 If angka0.Caption = "Angka" Then
 angka0.Caption = "Tunggu."


 ElseIf angka0.Caption = "Tunggu." Then
 angka0.Caption = "Tunggu.."
 sql = "select * from APOTEK where NMGenerik like '%" & resep2.nb5.Text & "%' order by NMGenerik asc"
 adopen.RecordSource = sql
 adopen.Recordset.Update
 adopen.Recordset.Fields(3) = angka2.Text
 adopen.Recordset.Update
 adopen.Refresh
 sql = " select * from APOTEK"
 adopen.RecordSource = sql
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.Refresh
  resep2.x.Enabled = False
selelaporan.Hide
 selelaporan.picket.Enabled = True
 selelaporan.sisaa2.Text = angka.Text
 selelaporan.outstok2.Text = angka.Text
 ElseIf angka0.Caption = "Tunggu.." Then
 angka0.Caption = "Tunggu..."

 ElseIf angka0.Caption = "Tunggu..." Then
 angka0.Caption = "Selesai"
 adopen.RecordSource = "APOTEK"
 adopen.Refresh
 Set dgpen.DataSource = adopen
 dgpen.AllowUpdate = False
 dgpen.TabStop = False
 dgpen.Refresh

 Unload Me
 Else
 MsgBox "FALSE PARSE !)"
 End If
End If
End Sub
