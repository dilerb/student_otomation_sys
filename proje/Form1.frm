VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Ders Programý Otomasyonu"
   ClientHeight    =   7650
   ClientLeft      =   2460
   ClientTop       =   1680
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Öðretmen Giriþi"
      Height          =   6735
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CommandButton Command5 
         Caption         =   "Düzenle"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7680
         TabIndex        =   29
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Yeni Kayýt"
         Height          =   615
         Left            =   5880
         TabIndex        =   28
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "LÝSTELE"
         Height          =   615
         Left            =   4080
         TabIndex        =   27
         Top             =   4440
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2655
         Left            =   3960
         TabIndex        =   26
         Top             =   1200
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   1
         Cols            =   5
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   405
         Index           =   5
         Left            =   1680
         TabIndex        =   25
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   405
         Index           =   4
         Left            =   1680
         TabIndex        =   24
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   405
         Index           =   3
         Left            =   1680
         TabIndex        =   23
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   405
         Index           =   2
         Left            =   1680
         TabIndex        =   22
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   405
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Telefon"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Adres"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Adý Soyadý"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Ünvan"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Sicil Numarasý"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sýnýf Giriþi"
      Height          =   6735
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   10815
      Begin VB.TextBox text3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   52
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox text3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   51
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox text3 
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   1200
         TabIndex        =   50
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox text3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   49
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo10 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   48
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox Combo9 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   47
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox Combo8 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   46
         Top             =   3360
         Width           =   615
      End
      Begin VB.ComboBox Combo7 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   45
         Top             =   2760
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ý.Ö"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   2160
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "N.Ö"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Düzenle"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7560
         TabIndex        =   42
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Yeni Kayýt"
         Height          =   615
         Left            =   5640
         TabIndex        =   41
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "LÝSTELE"
         Height          =   615
         Left            =   3600
         TabIndex        =   40
         Top             =   3600
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   2175
         Left            =   3360
         TabIndex        =   39
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   7
      End
      Begin VB.Label Label20 
         Caption         =   "Mevcut"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Danýþman"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Program"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Öðretim"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Þube"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Sýnýf"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Sýnýf Kodu"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Height          =   6735
      Left            =   360
      TabIndex        =   71
      Top             =   480
      Width           =   10815
      Begin VB.Timer Timer2 
         Left            =   1200
         Top             =   6000
      End
      Begin VB.Timer Timer1 
         Left            =   480
         Top             =   6000
      End
      Begin VB.Label Label21 
         Caption         =   "DERS PROGRAMI OTOMOSYONU                             "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   74
         Top             =   360
         Width           =   8055
      End
      Begin VB.Label Label23 
         Caption         =   "LÜTFEN SOL ÜSTTE BULUNAN VERÝ GÝRÝÞÝ SEKMESÝNE TIKLAYARAK VERÝ GÝRÝÞÝNÝ BAÞLATINIZ"
         Height          =   735
         Left            =   3840
         TabIndex        =   73
         Top             =   4560
         Width           =   3855
      End
      Begin VB.Label Label22 
         Caption         =   $"Form1.frx":0000
         Height          =   1935
         Left            =   3840
         TabIndex        =   72
         Top             =   2040
         Width           =   4095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sýnýf - Ders - Öðretmen Eþleþtirmesi"
      Enabled         =   0   'False
      Height          =   6735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   7215
      Begin VB.OptionButton Option1 
         Caption         =   "4+0"
         Height          =   495
         Index           =   3
         Left            =   4200
         TabIndex        =   15
         Top             =   3240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2+2"
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   14
         Top             =   2880
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2+0"
         Height          =   615
         Index           =   1
         Left            =   3720
         TabIndex        =   13
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         Height          =   2400
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Kaydet"
         Height          =   615
         Left            =   1200
         TabIndex        =   9
         Top             =   4560
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Text            =   "Derse Girecek Öðretmeni Seçiniz"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   840
         TabIndex        =   2
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Text            =   "Sýnýf Seçiniz"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Dersin saat bölümünü seçiniz"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Hangi sýnýfýn hangi dersi alacaðýný ve o derse hangi öðretmenin gireceðinin belirlenmesi"
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Þartlar"
      Enabled         =   0   'False
      Height          =   6735
      Left            =   7560
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ComboBox Combo11 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   70
         Top             =   720
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Onayla"
         Height          =   615
         Left            =   1200
         TabIndex        =   20
         Top             =   5280
         Width           =   1455
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Text            =   "Öðretmen Seçiniz"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3600
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label3 
         Caption         =   "Tatil Günü"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ders Giriþi"
      Height          =   6735
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   10815
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   1440
         TabIndex        =   38
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   37
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   35
         Top             =   1560
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2175
         Left            =   3600
         TabIndex        =   34
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   5
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Düzenle"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7920
         TabIndex        =   33
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Yeni Kayýt"
         Height          =   615
         Left            =   5760
         TabIndex        =   32
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "LÝSTELE"
         Height          =   615
         Left            =   3960
         TabIndex        =   31
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   30
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Uygulama"
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Teori"
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Kredi"
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Ders"
         Height          =   375
         Left            =   240
         TabIndex        =   59
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Ders Kodu"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Menu islem1 
      Caption         =   "Veri Giriþi"
      Begin VB.Menu verigir1 
         Caption         =   "Öðretmenler"
      End
      Begin VB.Menu verigir2 
         Caption         =   "Dersler"
      End
      Begin VB.Menu verigir3 
         Caption         =   "Sýnýflar"
      End
   End
   Begin VB.Menu islem2 
      Caption         =   "Eþleþtirme"
   End
   Begin VB.Menu islem3 
      Caption         =   "Ders Programýný Hazýrla!"
   End
   Begin VB.Menu islem4 
      Caption         =   "Raporlama"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
End Sub

Private Sub Combo1_Click()
Combo3.Text = Combo3.List(Combo1.ListIndex)
End Sub

Private Sub Combo10_Click()
Combo8.Text = Combo8.List(Combo10.ListIndex)
End Sub

Private Sub Combo2_Click()
Combo4.Text = Combo4.List(Combo2.ListIndex)
End Sub

Private Sub Combo5_Click()
Combo11.Text = Combo11.List(Combo5.ListIndex)
End Sub

Private Sub Combo9_click()
Combo7.Text = Combo7.List(Combo9.ListIndex)
End Sub

Private Sub Command1_Click()
Call vt_baglan
Call tbl_baglan("select*from ders_prog")
Call tbl2_baglan("select * from ders where ders_kodu='" & List2.List(List2.ListIndex) & "'")
If Combo1.Text = "Sýnýf Seçiniz" Then
MsgBox "Sýnýf Seçmediniz!"
GoTo bitis
End If
If List1.SelCount = 0 Then
MsgBox "Ders Seçmediniz!"
GoTo bitis
End If
If Combo2.Text = "Derse Girecek Öðretmeni Seçiniz" Then
MsgBox "Derse Girecek Olan Öðretmeni Seçmediniz!"
GoTo bitis
End If
If Option1(1).Value = False And Option1(2).Value = False And _
Option1(3).Value = False Then
MsgBox "Dersin saat bölümünü seçmediniz!"
GoTo bitis
End If
For i = 1 To 3
If Option1(i).Value = True Then
If (Val(Left(Option1(i).Caption, 1)) + Val(Right(Option1(i).Caption, 1))) <> (tbl2!teo + tbl2!uyg) Then
MsgBox "Bu dersin toplam saatiyle saat bölümü birbiriyle uymuyor!"
GoTo bitis
End If
End If
Next

Do While Not tbl.EOF
If Combo3.Text = tbl!sinif_kod Then
If List2.List(List2.ListIndex) = tbl!ders_kodu Then
MsgBox "Bu sýnýf ve ders daha önceden eþleþtirilmiþtir!"
GoTo bitis
End If
End If
tbl.MoveNext
Loop

tbl.AddNew
tbl!sinif_kod = Combo3.Text
tbl!ders_kodu = List2.List(List2.ListIndex)
tbl!sicil_no = Combo4.Text
For i = 1 To 3
If Option1(i).Value = True Then
tbl!saat_bol = Left(Option1(i).Caption, 1) & Right(Option1(i).Caption, 1)
Exit For
End If
Next
tbl.Update

MsgBox "Eþleþtirme Baþarýyla Kaydedildi"
veri_tab.Close
bitis:
End Sub

Private Sub Command10_Click()
If Command10.Caption = "Yeni Kayýt" Then
Call temizle3
MSFlexGrid3.Enabled = False
Command9.Enabled = False
Command11.Enabled = False
text3(1).Enabled = True
text3(2).Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Combo9.Enabled = True
Combo10.Enabled = True
text3(3).Enabled = True
Call vt_baglan
Call tbl_baglan("select * from sinif")
tbl.MoveLast
veri_tab.Close
Command10.Caption = "Kaydet"
Else

MSFlexGrid3.Enabled = True
Command9.Enabled = True
Command11.Enabled = True


Command10.Caption = "Yeni Kayýt"


Call kaydet3

End If
End Sub

Private Sub Command11_Click()
If Command11.Caption = "Düzenle" Then
MSFlexGrid3.Enabled = False
Command10.Enabled = False
Command9.Enabled = False
For i = 1 To 3
text3(i).Enabled = True
Next
Option2.Enabled = True
Option3.Enabled = True
Combo9.Enabled = True
Combo10.Enabled = True

Command11.Caption = "Deðiþliði Kaydet"
Else
MSFlexGrid3.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = False
For i = 1 To 3
text3(i).Enabled = False
Next
Call vt_baglan
Call tbl_baglan("select * from sinif where sinif_kod='" & text3(0).Text & "'")

tbl.Edit
tbl!sinif = Val(text3(1).Text)
tbl!sube = text3(2).Text
tbl!prog_kod = Combo7.Text
tbl!danisman = Combo8.Text
If Option2.Value = True Then tbl!ogretim = -1
If Option3.Value = True Then tbl!ogretim = 0
tbl!mevcut = Val(text3(3).Text)

tbl.Update
veri_tab.Close
Command11.Caption = "Düzenle"
End If
End Sub

Private Sub Command2_Click()
Call vt_baglan
Call tbl_baglan("select * from ogretmen where sicil_no='" & Combo11.Text & "'")
If Combo11.Text = "" Then
MsgBox "Öðretmen Seçmediniz!"
GoTo e1
End If
If Combo6.Text = "" Then
MsgBox "Gün Seçmediniz!"
GoTo e1
End If
tbl.Edit
tbl!sart = Combo6.Text
tbl.Update
veri_tab.Close
MsgBox "Þartýnýz Onaylanmýþtýr"
e1:
End Sub

Private Sub Command3_Click()
Command5.Enabled = True
MSFlexGrid1.Rows = 1
Call vt_baglan
Call tbl_baglan("select * from ogretmen")
Do While Not tbl.EOF
MSFlexGrid1.AddItem tbl!sicil_no & Chr(9) & tbl!unvan & Chr(9) & tbl!ad_soyad & Chr(9) & tbl!adres & Chr(9) & tbl!tel
tbl.MoveNext
Loop
veri_tab.Close
Text1(1).Enabled = False
Text1(2).Enabled = False
Text1(3).Enabled = False
Text1(4).Enabled = False
Text1(5).Enabled = False
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Yeni Kayýt" Then
Call temizle
MSFlexGrid1.Enabled = True
Command3.Enabled = False
Command5.Enabled = False
Text1(1).Enabled = True
Text1(2).Enabled = True
Text1(3).Enabled = True
Text1(4).Enabled = True
Text1(5).Enabled = True
Call vt_baglan
Call tbl_baglan("select * from ogretmen")
tbl.MoveLast
veri_tab.Close
Command4.Caption = "Kaydet"
Else

MSFlexGrid1.Enabled = True
Command3.Enabled = True
Command5.Enabled = True


Command4.Caption = "Yeni Kayýt"


Call kaydet

End If
End Sub

Private Sub Command5_Click()
If Command5.Caption = "Düzenle" Then
MSFlexGrid1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
For i = 2 To 5
Text1(i).Enabled = True
Next
Command5.Caption = "Deðiþliði Kaydet"
Else
MSFlexGrid1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
For i = 2 To 5
Text1(i).Enabled = False
Next
Call vt_baglan
Call tbl_baglan("select * from ogretmen where sicil_no='" & Text1(1).Text & "'")
tbl.Edit
tbl!unvan = Text1(2).Text
tbl!ad_soyad = Text1(3).Text
tbl!adres = Text1(4).Text
tbl!tel = Text1(5).Text
tbl.Update
veri_tab.Close
Command5.Caption = "Düzenle"
End If
End Sub

Private Sub Command6_Click()
Command8.Enabled = True
MSFlexGrid2.Rows = 1
Call vt_baglan
Call tbl_baglan("select * from ders")
Do While Not tbl.EOF
MSFlexGrid2.AddItem tbl!ders_kodu & Chr(9) & tbl!ders_adi & Chr(9) & tbl!kredi & Chr(9) & tbl!teo & Chr(9) & tbl!uyg
tbl.MoveNext
Loop
veri_tab.Close
Text2(1).Enabled = False
Text2(2).Enabled = False
Text2(3).Enabled = False
Text2(4).Enabled = False
Text2(5).Enabled = False
End Sub

Private Sub Command7_Click()
If Command7.Caption = "Yeni Kayýt" Then
Call temizle2
MSFlexGrid2.Enabled = True
Command6.Enabled = False
Command8.Enabled = False
Text2(1).Enabled = True
Text2(2).Enabled = True
Text2(3).Enabled = True
Text2(4).Enabled = True
Text2(5).Enabled = True
Call vt_baglan
Call tbl_baglan("select * from ders")
tbl.MoveLast
veri_tab.Close
Command7.Caption = "Kaydet"
Else

MSFlexGrid2.Enabled = True
Command6.Enabled = True
Command8.Enabled = True


Command7.Caption = "Yeni Kayýt"


Call kaydet2
End If
End Sub

Private Sub Command8_Click()
If Command8.Caption = "Düzenle" Then


MSFlexGrid2.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
For i = 2 To 5
Text2(i).Enabled = True
Next
Command8.Caption = "Deðiþliði Kaydet"

Else
MSFlexGrid2.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
For i = 2 To 5
Text2(i).Enabled = False

Next

Call vt_baglan
Call tbl_baglan("select * from ders where ders_kodu='" & Text2(1).Text & "'")
tbl.Edit
tbl!ders_adi = Text2(2).Text
tbl!kredi = Text2(3).Text
tbl!teo = Text2(4).Text
tbl!uyg = Text2(5).Text
tbl.Update
veri_tab.Close
Command8.Caption = "Düzenle"
End If
End Sub

Private Sub Command9_Click()
Command11.Enabled = True
MSFlexGrid3.Rows = 1
Call vt_baglan
Call tbl_baglan("select * from sinif order by sinif_kod asc")

Do While Not tbl.EOF
Call tbl3_baglan("select * from program where prog_kod=" & tbl!prog_kod)
If tbl3.EOF Then
program = ""
Combo7.Text = ""
Else
program = tbl3!prog_adi
End If
Call tbl4_baglan("select * from ogretmen where sicil_no='" & tbl!danisman & "'")
If tbl4.EOF Then
danisman = ""
Combo8.Text = ""
Else
danisman = tbl4!ad_soyad
End If
If tbl!ogretim = 0 Then
ogretim = "Ý.Ö"
Else
ogretim = "N.Ö"
End If
MSFlexGrid3.AddItem tbl!sinif_kod & Chr(9) & tbl!sinif & Chr(9) & tbl!sube & Chr(9) & ogretim & Chr(9) & program & Chr(9) & danisman & Chr(9) & tbl!mevcut
tbl.MoveNext
Loop
veri_tab.Close
text3(1).Enabled = False
text3(2).Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Combo7.Enabled = False
Combo8.Enabled = False
text3(3).Enabled = False
End Sub

Private Sub Form_Load()
Frame6.Visible = True
Timer1.Interval = 150
End Sub

Private Sub islem2_Click()
Frame5.Enabled = True
Frame5.Visible = True
Frame6.Visible = False
islem2.Enabled = False
Frame4.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Call vt_baglan
Call tbl_baglan("select * from sinif")
Call tbl2_baglan("select * from program")
Call tbl3_baglan("select * from ders order by ders_adi asc")
Call tbl4_baglan("select * from ogretmen order by ad_soyad asc")
Do While Not tbl.EOF
If tbl!ogretim = -1 Then ogrt = "(N.Ö)"
If tbl!ogretim = 0 Then ogrt = "(Ý.Ö)"
Combo3.AddItem tbl!sinif_kod
Combo1.AddItem tbl2!prog_adi & " " & tbl!sinif & tbl!sube & " " & ogrt
tbl.MoveNext
Loop
Do While Not tbl3.EOF
List1.AddItem tbl3!ders_adi
List2.AddItem tbl3!ders_kodu
tbl3.MoveNext
Loop
Do While Not tbl4.EOF
Combo2.AddItem tbl4!ad_soyad
Combo5.AddItem tbl4!ad_soyad
Combo4.AddItem tbl4!sicil_no
Combo11.AddItem tbl4!sicil_no
tbl4.MoveNext
Loop
Combo6.AddItem "Pazartesi"
Combo6.AddItem "Salý"
Combo6.AddItem "Çarþamba"
Combo6.AddItem "Perþembe"
Combo6.AddItem "Cuma"
veri_tab.Close
End Sub

Private Sub islem3_Click()
Unload Me
Form2.Show
End Sub

Private Sub islem4_Click()
Unload Me
Form3.Show
End Sub

Private Sub List1_Click()
List2.Selected(List1.ListIndex) = True
End Sub

Private Sub MSFlexGrid1_dblClick()
On Local Error GoTo hata
sicilno = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
Call vt_baglan
Call tbl_baglan("select * from ogretmen where sicil_no='" & sicilno & "'")
musterino = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
Text1(1).Text = tbl!sicil_no
Text1(2).Text = tbl!unvan
Text1(3).Text = tbl!ad_soyad
Text1(4).Text = tbl!adres
Text1(5).Text = tbl!tel
Exit Sub
hata:
If Err = 94 Then Exit Sub
veri_tab.Close
End Sub

Private Sub MSFlexGrid2_dblClick()
On Local Error GoTo hata
derskodu = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0)
Call vt_baglan
Call tbl_baglan("select * from ders where ders_kodu='" & derskodu & "'")
sicilno = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0)
Text2(1).Text = tbl!ders_kodu
Text2(2).Text = tbl!ders_adi
Text2(3).Text = tbl!kredi
Text2(4).Text = tbl!teo
Text2(5).Text = tbl!uyg
Exit Sub
hata:
If Err = 94 Then Exit Sub
veri_tab.Close
End Sub

Private Sub MSFlexGrid3_dblClick()
sinif = MSFlexGrid3.TextMatrix(MSFlexGrid3.Row, 0)
Call vt_baglan
Call tbl_baglan("select * from sinif where sinif_kod='" & sinif & "'")
text3(0).Text = tbl!sinif_kod
text3(1).Text = tbl!sinif
text3(2).Text = tbl!sube
If tbl!ogretim = -1 Then
Option2.Value = 1
Option3.Value = 0
End If
If tbl!ogretim = 0 Then
Option2.Value = 0
Option3.Value = 1
End If
Combo7.Text = tbl!prog_kod
text3(3).Text = tbl!mevcut

tbl.MoveFirst
Call tbl_baglan("select * from sinif where sinif_kod='" & text3(0).Text & "'")
Combo7.Text = tbl!prog_kod
Combo8.Text = tbl!danisman
Call tbl3_baglan("select * from program where prog_kod=" & Combo7.Text)
If tbl3.EOF Then
Combo9.Text = ""
Combo7.Text = ""
Else
Combo9.Text = tbl3!prog_adi
End If
Call tbl4_baglan("select * from ogretmen where sicil_no='" & Combo8.Text & "'")
If tbl4.EOF Then
Combo10.Text = ""
Combo8.Text = ""
Else
Combo10.Text = tbl4!ad_soyad
End If
veri_tab.Close
End Sub

Private Sub Timer1_Timer()
Label21.Caption = Mid(Label21.Caption, 2) + Left(Label21.Caption, 1)
End Sub

Private Sub Timer2_Timer()
If Label23.Visible = True Then

Label23.Visible = False

Else

Label23.Visible = True
End If
End Sub

Private Sub verigir1_Click()
islem2.Enabled = True
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
MSFlexGrid1.TextMatrix(0, 0) = "Sicil No"
MSFlexGrid1.TextMatrix(0, 1) = "Ünvan"
MSFlexGrid1.TextMatrix(0, 2) = "Ad Soyad"
MSFlexGrid1.TextMatrix(0, 3) = "Adres"
MSFlexGrid1.TextMatrix(0, 4) = "Telefon"
End Sub

Private Sub verigir2_Click()
islem2.Enabled = True
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
MSFlexGrid2.TextMatrix(0, 0) = "Ders Kodu"
MSFlexGrid2.TextMatrix(0, 1) = "Ders"
MSFlexGrid2.TextMatrix(0, 2) = "Kredi"
MSFlexGrid2.TextMatrix(0, 3) = "Teori"
MSFlexGrid2.TextMatrix(0, 4) = "Uygulama"
End Sub

Private Sub verigir3_Click()
islem2.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
MSFlexGrid3.TextMatrix(0, 0) = "Sýnýf Kodu"
MSFlexGrid3.TextMatrix(0, 1) = "Sýnýf"
MSFlexGrid3.TextMatrix(0, 2) = "Þube"
MSFlexGrid3.TextMatrix(0, 3) = "Öðretim"
MSFlexGrid3.TextMatrix(0, 4) = "program"
MSFlexGrid3.TextMatrix(0, 5) = "Danýþman"
MSFlexGrid3.TextMatrix(0, 6) = "mevcut"
Call vt_baglan
Call tbl_baglan("select * from program ")
Call tbl2_baglan("select * from ogretmen")
Do While Not tbl.EOF
Combo7.AddItem tbl!prog_kod
Combo9.AddItem tbl!prog_adi
tbl.MoveNext
Loop
Do While Not tbl2.EOF
Combo8.AddItem tbl2!sicil_no
Combo10.AddItem tbl2!ad_soyad
tbl2.MoveNext
Loop
veri_tab.Close
End Sub

Sub temizle()
For i = 1 To 5
Text1(i).Text = ""
Next
End Sub

Sub kaydet()
On Local Error GoTo hata2
Call vt_baglan
Call tbl_baglan("select * from ogretmen")
If Text1(1).Text = "" Then
MsgBox "sicil numarasý boþ geçilemez"
GoTo e1
End If
tbl.AddNew
tbl!sicil_no = Text1(1).Text
tbl!unvan = Text1(2).Text
tbl!ad_soyad = Text1(3).Text
tbl!adres = Text1(4).Text
tbl!tel = Text1(5).Text
tbl.Update
veri_tab.Close
Exit Sub
hata2:
If Err = 381 Then Exit Sub
e1:
End Sub

Sub kaydet2()
On Local Error GoTo hata2
Call vt_baglan
Call tbl_baglan("select * from ders")
If Text2(1).Text = "" Then
MsgBox "ders kodu boþ geçilemez"
GoTo e1
End If
tbl.AddNew
tbl!ders_kodu = Text2(1).Text
tbl!ders_adi = Text2(2).Text
tbl!kredi = Text2(3).Text
tbl!teo = Text2(4).Text
tbl!uyg = Text2(5).Text
tbl.Update
veri_tab.Close
Exit Sub
hata2:
If Err = 381 Then Exit Sub
e1:
End Sub

Sub temizle3()
For i = 1 To 3
text3(i).Text = ""
Next
Option2.Value = False
Option3.Value = False
Combo9.Text = ""
Combo10.Text = ""
End Sub

Sub kaydet3()
Call vt_baglan
Call tbl_baglan("select * from sinif")
Call tbl2_baglan("select count(*) as i from sinif")
tbl.AddNew

tbl!sinif_kod = tbl2!i + 1
tbl!sinif = Val(text3(1).Text)
tbl!sube = text3(2).Text
tbl!prog_kod = Val(Combo7.Text)
tbl!danisman = Val(Combo8.Text)
If Option2.Value = True Then tbl!ogretim = -1
If Option3.Value = True Then tbl!ogretim = 0

tbl!mevcut = Val(text3(3).Text)
tbl.Update
veri_tab.Close
End Sub


Sub temizle2()
For i = 1 To 5
Text2(i).Text = ""
Next
End Sub
