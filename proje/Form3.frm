VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8130
   ClientLeft      =   2460
   ClientTop       =   420
   ClientWidth     =   12030
   LinkTopic       =   "Form3"
   ScaleHeight     =   8130
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Listele"
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      Height          =   2985
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   4680
      TabIndex        =   0
      Top             =   2640
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   1
      Cols            =   6
   End
   Begin VB.Label Label1 
      Caption         =   "Sýnýf Seçiniz"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call vt_baglan
Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "'")
'Call tbl4_baglan("select * from sinif where sinif_kod='" & List4.List(List1.ListIndex) & "'")
Do While Not tbl.EOF
Select Case tbl!saat_bol

Case 20
Call tbl2_baglan("select * from ders where ders_kodu='" & tbl!ders_kodu & "'")
If tbl!gun = "Pazartesi" Then t2 = 1
If tbl!gun = "Salý" Then t2 = 2
If tbl!gun = "Çarþamba" Then t2 = 3
If tbl!gun = "Perþembe" Then t2 = 4
If tbl!gun = "Cuma" Then t2 = 5
t1 = tbl!saat
MSFlexGrid1.TextMatrix(t1, t2) = tbl2!ders_adi
MSFlexGrid1.TextMatrix(t1 + 1, t2) = tbl2!ders_adi
Case 40
Call tbl2_baglan("select * from ders where ders_kodu='" & tbl!ders_kodu & "'")
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Pazartesi" Then t2 = 1
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Salý" Then t2 = 2
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Çarþamba" Then t2 = 3
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Perþembe" Then t2 = 4
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Cuma" Then t2 = 5
t1 = Val((Left(tbl!saat, InStr(1, tbl!saat, ",") - 1)))
MSFlexGrid1.TextMatrix(t1, t2) = tbl2!ders_adi
MSFlexGrid1.TextMatrix(t1 + 1, t2) = tbl2!ders_adi
MSFlexGrid1.TextMatrix(t1 + 2, t2) = tbl2!ders_adi
MSFlexGrid1.TextMatrix(t1 + 3, t2) = tbl2!ders_adi

Case 22
Call tbl2_baglan("select * from ders where ders_kodu='" & tbl!ders_kodu & "'")
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Pazartesi" Then t2 = 1
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Salý" Then t2 = 2
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Çarþamba" Then t2 = 3
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Perþembe" Then t2 = 4
If Left(tbl!gun, InStr(1, tbl!gun, ",") - 1) = "Cuma" Then t2 = 5
t1 = Val((Left(tbl!saat, InStr(1, tbl!saat, ",") - 1)))
MSFlexGrid1.TextMatrix(t1, t2) = tbl2!ders_adi
MSFlexGrid1.TextMatrix(t1 + 1, t2) = tbl2!ders_adi

If Right(tbl!gun, Len(tbl!gun) - InStr(1, tbl!gun, ",")) = "Pazartesi" Then t2 = 1
If Right(tbl!gun, Len(tbl!gun) - InStr(1, tbl!gun, ",")) = "Salý" Then t2 = 2
If Right(tbl!gun, Len(tbl!gun) - InStr(1, tbl!gun, ",")) = "Çarþamba" Then t2 = 3
If Right(tbl!gun, Len(tbl!gun) - InStr(1, tbl!gun, ",")) = "Perþembe" Then t2 = 4
If Right(tbl!gun, Len(tbl!gun) - InStr(1, tbl!gun, ",")) = "Cuma" Then t2 = 5

t1 = Val((Right(tbl!saat, Len(tbl!saat) - InStr(1, tbl!saat, ","))))
MSFlexGrid1.TextMatrix(t1, t2) = tbl2!ders_adi
MSFlexGrid1.TextMatrix(t1 + 1, t2) = tbl2!ders_adi


End Select

tbl.MoveNext
Loop
End Sub

Private Sub Form_Load()
Call vt_baglan
Call tbl_baglan("select * from saat")
Call tbl2_baglan("select * from ogretmen")
Call tbl4_baglan("select * from ders")
Call tbl5_baglan("select * from sinif ")
Call tbl7_baglan("select * from program")

Do While Not tbl5.EOF
If tbl5!ogretim = -1 Then ogrt = "(N.Ö)"
If tbl5!ogretim = 0 Then ogrt = "(Ý.Ö)"
List4.AddItem tbl5!sinif_kod
List1.AddItem tbl7!prog_adi & " " & tbl5!sinif & tbl5!sube & " " & ogrt
tbl5.MoveNext
Loop

MSFlexGrid1.TextMatrix(0, 1) = "Pazartesi"
MSFlexGrid1.TextMatrix(0, 2) = "Salý"
MSFlexGrid1.TextMatrix(0, 3) = "Çarþamba"
MSFlexGrid1.TextMatrix(0, 4) = "Perþembe"
MSFlexGrid1.TextMatrix(0, 5) = "Cuma"

Do While Not tbl.EOF
MSFlexGrid1.AddItem tbl!ders_saati
tbl.MoveNext
Loop

veri_tab.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Form1.Show
End Sub

Private Sub List1_Click()
dongu = 0
Label1.Caption = List1.List(List1.ListIndex) & " sýnýfýnýn aldýðý dersler"
Command1.Enabled = True
For s = 1 To 5
For i = 1 To 12
MSFlexGrid1.TextMatrix(i, s) = ""
Next
Next

Call vt_baglan
Call tbl_baglan("select * from ders")
Call tbl3_baglan("select * from ogretmen")
Call tbl2_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "'")
Call tbl4_baglan("select * from sinif where sinif_kod='" & List4.List(List1.ListIndex) & "'")
Do While Not tbl2.EOF
Call tbl_baglan("select * from ders where ders_kodu='" & tbl2!ders_kodu & "'")
Call tbl3_baglan("select * from ogretmen where sicil_no='" & tbl2!sicil_no & "'")

tbl2.MoveNext
Loop
If tbl4!ogretim = -1 Then
For s = 1 To 5
For i = 7 To 12
MSFlexGrid1.TextMatrix(i, s) = "X"
Next
Next
Else
For s = 1 To 5
For i = 1 To 6
MSFlexGrid1.TextMatrix(i, s) = "X"
Next
Next
End If
veri_tab.Close
End Sub

