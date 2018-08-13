VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Ders Programýnýn Hazýrlanmasý"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13140
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   1935
      Left            =   9360
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      Enabled         =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1695
      Left            =   9120
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      Enabled         =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hazýrla"
      Enabled         =   0   'False
      Height          =   855
      Left            =   7680
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   9120
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   5040
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   5760
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3255
      Left            =   1080
      TabIndex        =   0
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      MouseIcon       =   "Form2.frx":0152
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bas_flex_x As Long
Dim bas_flex_y As Long
Dim hedef_flex_x As Long
Dim hedef_flex_y As Long
Dim c1 As Variant
Dim c2 As Variant
Dim c3 As Variant
Dim dongu As Integer

Private Sub Command3_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command1_Click()
e2:
Dim t1, t2, ilkt1, ilkt2 As Integer
Call vt_baglan
Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "' order by saat_bol desc")
Call tbl4_baglan("select * from sinif where sinif_kod='" & List4.List(List1.ListIndex) & "'")

If tbl!gun <> "" Then
MsgBox "Bu sýnýf için ders programý önceden hazýrlanmýþtýr!"
GoTo bitis2
End If

If Command1.Caption = "Hazýrla" Then
Randomize Timer

Do While Not tbl.EOF
basadon:
'''''''''''''''''''''''''''''''''''''döngü
dongu = dongu + 1
If dongu = 1000 Then

dongu = 0
For s = 1 To 5
For i = 1 To 12
MSFlexGrid1.TextMatrix(i, s) = ""
Next
Next
For s = 1 To 5
For i = 1 To 12
MSFlexGrid2.TextMatrix(i, s) = ""
Next
Next
If tbl4!ogretim = -1 Then
For s = 1 To 5
For i = 7 To 12
MSFlexGrid1.TextMatrix(i, s) = "X"
MSFlexGrid2.TextMatrix(i, s) = "X"
Next
Next
Else
For s = 1 To 5
For i = 1 To 6
MSFlexGrid1.TextMatrix(i, s) = "X"
MSFlexGrid2.TextMatrix(i, s) = "X"
Next
Next
End If
veri_tab.Close
GoTo e2
End If
''''''''''''''''''''''''''''''''''''döngü
If tbl4!ogretim = -1 Then
Select Case tbl!saat_bol

Case 40

tekrar2:
t1 = Int(Rnd * 3 + 1)
If t1 = 2 Then GoTo tekrar2
t2 = Int(Rnd * 5 + 1)

Case 22

tekrar3:
t1 = Int(Rnd * 5 + 1)
If t1 = 2 Or t1 = 4 Then GoTo tekrar3
t2 = Int(Rnd * 5 + 1)
tekrar8:
t3 = Int(Rnd * 5 + 1)
If t3 = 2 Or t3 = 4 Then GoTo tekrar8
t4 = Int(Rnd * 5 + 1)
If t2 = t4 Then
If t1 = 1 And t3 = 3 Then GoTo tekrar8
If t1 = 3 And t3 = 5 Then GoTo tekrar8
End If

Case 20

tekrar4:
t1 = Int(Rnd * 5 + 1)
If t1 = 2 Or t1 = 4 Then GoTo tekrar4
t2 = Int(Rnd * 5 + 1)
End Select

Else

Select Case tbl!saat_bol

Case 40

tekrar5:
t1 = Int(Rnd * 3 + 7)
If t1 = 8 Or t1 = 10 Or t1 = 11 Then GoTo tekrar5
t2 = Int(Rnd * 5 + 1)

Case 22

tekrar6:
t1 = Int(Rnd * 5 + 7)
If t1 = 8 Or t1 = 10 Then GoTo tekrar6
t2 = Int(Rnd * 5 + 1)
ilkt1 = t1
ilkt2 = t2

Case 20

tekrar7:
t1 = Int(Rnd * 5 + 7)
If t1 = 8 Or t1 = 10 Then GoTo tekrar7
t2 = Int(Rnd * 5 + 1)
End Select
End If

If MSFlexGrid2.TextMatrix(t1, t2) <> "" Then GoTo basadon


'''''''''''''''''''''''''''''''''''''''''''
'çakýþmalar...

If t2 = 1 Then gun = "Pazartesi"
If t2 = 2 Then gun = "Salý"
If t2 = 3 Then gun = "Çarþamba"
If t2 = 4 Then gun = "Perþembe"
If t2 = 5 Then gun = "Cuma"

Select Case tbl!saat_bol

Case 20
Call tbl5_baglan("select * from ders_prog where sinif_kod<>'" & List4.List(List1.ListIndex) & "' and sicil_no='" & tbl!sicil_no & "'")
Call tbl2_baglan("select * from ogretmen where sicil_no='" & tbl!sicil_no & "'")
If tbl2!sart <> "" Then
If gun = tbl2!sart Then GoTo basadon
End If
Do While Not tbl5.EOF
If tbl5!saat_bol = 20 Then
If t1 = tbl5!saat And tbl5!gun = gun Then GoTo basadon
Else
If t1 = Val((Left(tbl5!saat, InStr(1, tbl5!saat, ",") - 1))) And gun = Left(tbl5!gun, InStr(1, tbl5!gun, ",") - 1) Then GoTo basadon
If t1 = Val((Right(tbl5!saat, Len(tbl5!saat) - InStr(1, tbl5!saat, ",")))) And gun = Right(tbl5!gun, Len(tbl5!gun) - InStr(1, tbl5!gun, ",")) Then GoTo basadon
End If
tbl5.MoveNext
Loop

Case 40
Call tbl5_baglan("select * from ders_prog where sinif_kod<>'" & List4.List(List1.ListIndex) & "' and sicil_no='" & tbl!sicil_no & "'")
Call tbl2_baglan("select * from ogretmen where sicil_no='" & tbl!sicil_no & "'")
If tbl2!sart <> "" Then
If gun = tbl2!sart Then GoTo basadon
End If
Do While Not tbl5.EOF
If tbl5!saat_bol = 20 Then
If t1 = tbl5!saat And tbl5!gun = gun Then GoTo basadon
If t1 + 2 = tbl5!saat And tbl5!gun = gun Then GoTo basadon
Else
If t1 = Val((Left(tbl5!saat, InStr(1, tbl5!saat, ",") - 1))) And gun = Left(tbl5!gun, InStr(1, tbl5!gun, ",") - 1) Then GoTo basadon
If t1 = Val((Right(tbl5!saat, Len(tbl5!saat) - InStr(1, tbl5!saat, ",")))) And gun = Right(tbl5!gun, Len(tbl5!gun) - InStr(1, tbl5!gun, ",")) Then GoTo basadon
If t1 + 2 = Val((Left(tbl5!saat, InStr(1, tbl5!saat, ",") - 1))) And gun = Left(tbl5!gun, InStr(1, tbl5!gun, ",") - 1) Then GoTo basadon
If t1 + 2 = Val((Right(tbl5!saat, Len(tbl5!saat) - InStr(1, tbl5!saat, ",")))) And gun = Right(tbl5!gun, Len(tbl5!gun) - InStr(1, tbl5!gun, ",")) Then GoTo basadon
End If
tbl5.MoveNext
Loop

Case 22

Call tbl5_baglan("select * from ders_prog where sinif_kod<>'" & List4.List(List1.ListIndex) & "' and sicil_no='" & tbl!sicil_no & "'")
Call tbl2_baglan("select * from ogretmen where sicil_no='" & tbl!sicil_no & "'")
If tbl2!sart <> "" Then
If gun = tbl2!sart Then GoTo basadon
End If
Do While Not tbl5.EOF
If tbl5!saat_bol = 20 Then
If t1 = tbl5!saat And tbl5!gun = gun Then GoTo basadon
Else
If t1 = Val((Left(tbl5!saat, InStr(1, tbl5!saat, ",") - 1))) And gun = Left(tbl5!gun, InStr(1, tbl5!gun, ",") - 1) Then GoTo basadon
If t1 = Val((Right(tbl5!saat, Len(tbl5!saat) - InStr(1, tbl5!saat, ",")))) And gun = Right(tbl5!gun, Len(tbl5!gun) - InStr(1, tbl5!gun, ",")) Then GoTo basadon
End If
tbl5.MoveNext
Loop

End Select
'''''''''''''''''''''''''''''''''''''''''''

Call tbl3_baglan("select * from ders where ders_kodu='" & tbl!ders_kodu & "'")
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu

Select Case tbl!saat_bol
Case 40
t1 = t1 + 1
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu
t1 = t1 + 1
If MSFlexGrid1.TextMatrix(t1, t2) <> "" Then
MSFlexGrid1.TextMatrix(t1 - 1, t2) = ""
MSFlexGrid2.TextMatrix(t1 - 1, t2) = ""
MSFlexGrid1.TextMatrix(t1 - 2, t2) = ""
MSFlexGrid2.TextMatrix(t1 - 2, t2) = ""
GoTo basadon
End If
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu
t1 = t1 + 1
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu

Case 22

t1 = t1 + 1
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu
tekrar9:
t1 = Int(Rnd * 5 + 7)
If t1 = 8 Or t1 = 10 Then GoTo tekrar9
t2 = Int(Rnd * 5 + 1)
If t1 = ilkt1 + 2 And t2 = ilkt2 Then GoTo tekrar9
If t1 = ilkt1 - 2 And t2 = ilkt2 Then GoTo tekrar9
If MSFlexGrid1.TextMatrix(t1, t2) <> "" Then GoTo tekrar9
'''
'çakýþma
If t2 = 1 Then gun = "Pazartesi"
If t2 = 2 Then gun = "Salý"
If t2 = 3 Then gun = "Çarþamba"
If t2 = 4 Then gun = "Perþembe"
If t2 = 5 Then gun = "Cuma"
'''
Call tbl5_baglan("select * from ders_prog where sinif_kod<>'" & List4.List(List1.ListIndex) & "' and sicil_no='" & tbl!sicil_no & "'")
Call tbl2_baglan("select * from ogretmen where sicil_no='" & tbl!sicil_no & "'")
If tbl2!sart <> "" Then
If gun = tbl2!sart Then GoTo tekrar9
End If
'If tbl5.EOF Then
'GoTo e2
'Else
'tbl5.MoveNext
'End If
Do While Not tbl5.EOF
If tbl5!saat_bol = 20 Then
If t1 = tbl5!saat And tbl5!gun = gun Then GoTo tekrar9
Else
If t1 = Val((Left(tbl5!saat, InStr(1, tbl5!saat, ",") - 1))) And gun = Left(tbl5!gun, InStr(1, tbl5!gun, ",") - 1) Then GoTo tekrar9
If t1 = Val((Right(tbl5!saat, Len(tbl5!saat) - InStr(1, tbl5!saat, ",")))) And gun = Right(tbl5!gun, Len(tbl5!gun) - InStr(1, tbl5!gun, ",")) Then GoTo tekrar9
End If
tbl5.MoveNext
Loop
'''
'e2:
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu
t1 = t1 + 1
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu

Case 20

t1 = t1 + 1
MSFlexGrid1.TextMatrix(t1, t2) = tbl3!ders_adi
MSFlexGrid2.TextMatrix(t1, t2) = tbl3!ders_kodu
End Select

tbl.MoveNext

Loop
veri_tab.Close
Command1.Caption = "Kaydet"
Else
Command1.Caption = "Hazýrla"
'
''''KAYDETME ÝÞLEMLERÝ
If tbl4!ogretim = -1 Then
'Normal
'Öðretim
'Ýçin

i = 1
s = 1
Do While s < 6
Do While i < 6
e6:
If MSFlexGrid2.TextMatrix(i, s) = "" Then
i = i + 1
If i = 6 Then
i = 1
s = s + 1
If s = 6 Then GoTo e9
End If
GoTo e6
End If

Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "' and ders_kodu='" & MSFlexGrid2.TextMatrix(i, s) & "'")
Select Case tbl!saat_bol
Case 20
tbl.Edit
tbl!saat = i
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
tbl!gun = gun
tbl.Update
i = i + 1
Case 40
tbl.Edit
tbl!saat = i & "," & (i + 2)
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
tbl!gun = gun & "," & gun
i = i + 3
tbl.Update

Case 22
tbl.Edit
tbl!saat = i
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
tbl!gun = gun
tbl.Update
For t2 = 1 To 5
For t1 = 1 To 6 Step 2
If MSFlexGrid2.TextMatrix(t1, t2) = MSFlexGrid2.TextMatrix(i, s) Then
If t1 = i And t2 = s Then GoTo e8
If t1 <> i Or t2 <> s Then
tbl.Edit
tbl!saat = i & "," & t1
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
If t2 = 1 Then gun2 = "Pazartesi"
If t2 = 2 Then gun2 = "Salý"
If t2 = 3 Then gun2 = "Çarþamba"
If t2 = 4 Then gun2 = "Perþembe"
If t2 = 5 Then gun2 = "Cuma"
tbl!gun = gun & "," & gun2
tbl.Update
i = i + 1
GoTo e7
End If
End If
e8:
Next
Next

End Select
e7:
i = i + 1
Loop
i = 1
s = s + 1
Loop
e9:
'
'
Call salon_kaydet
MsgBox "Ders programý baþarýyla kaydedildi"
veri_tab.Close

Else
''''Ý.Ö ÝÇÝN
'KAYDETME

i = 7
s = 1
Do While s < 6
Do While i < 12
e1:
If MSFlexGrid2.TextMatrix(i, s) = "" Then
i = i + 1
If i = 12 Then
i = 7
s = s + 1
If s = 6 Then GoTo e5
End If
GoTo e1
End If

Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "' and ders_kodu='" & MSFlexGrid2.TextMatrix(i, s) & "'")
Select Case tbl!saat_bol
Case 20
tbl.Edit
tbl!saat = i
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
tbl!gun = gun
tbl.Update
i = i + 1
Case 40
tbl.Edit
tbl!saat = i & "," & (i + 2)
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
tbl!gun = gun & "," & gun
i = i + 3
tbl.Update

Case 22
tbl.Edit
tbl!saat = i
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
tbl!gun = gun
tbl.Update
For t2 = 1 To 5
For t1 = 7 To 12 Step 2
If MSFlexGrid2.TextMatrix(t1, t2) = MSFlexGrid2.TextMatrix(i, s) Then
If t1 = i And t2 = s Then GoTo e3
If t1 <> i Or t2 <> s Then
tbl.Edit
tbl!saat = i & "," & t1
If s = 1 Then gun = "Pazartesi"
If s = 2 Then gun = "Salý"
If s = 3 Then gun = "Çarþamba"
If s = 4 Then gun = "Perþembe"
If s = 5 Then gun = "Cuma"
If t2 = 1 Then gun2 = "Pazartesi"
If t2 = 2 Then gun2 = "Salý"
If t2 = 3 Then gun2 = "Çarþamba"
If t2 = 4 Then gun2 = "Perþembe"
If t2 = 5 Then gun2 = "Cuma"
tbl!gun = gun & "," & gun2
tbl.Update
i = i + 1
GoTo e4
End If
End If
e3:
Next
Next

End Select
e4:
i = i + 1
Loop
i = 7
s = s + 1
Loop
End If
e5:
'
'
Call salon_kaydet2
MsgBox "Ders programý baþarýyla kaydedildi"
veri_tab.Close

End If
bitis2:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Form1.Show
End Sub

Private Sub Frame1_click()
Timer1.Enabled = True
End Sub

Private Sub List1_Click()
dongu = 0
Label1.Caption = List1.List(List1.ListIndex) & " sýnýfýnýn aldýðý dersler"
Command1.Enabled = True
Command1.Caption = "Hazýrla"
For s = 1 To 5
For i = 1 To 12
MSFlexGrid1.TextMatrix(i, s) = ""
MSFlexGrid2.TextMatrix(i, s) = ""
MSFlexGrid3.TextMatrix(i, s) = ""
Next
Next

List2.Clear
List3.Clear
List5.Clear
Call vt_baglan
Call tbl_baglan("select * from ders")
Call tbl3_baglan("select * from ogretmen")
Call tbl2_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "'")
Call tbl4_baglan("select * from sinif where sinif_kod='" & List4.List(List1.ListIndex) & "'")
Do While Not tbl2.EOF
Call tbl_baglan("select * from ders where ders_kodu='" & tbl2!ders_kodu & "'")
Call tbl3_baglan("select * from ogretmen where sicil_no='" & tbl2!sicil_no & "'")
List3.AddItem tbl!ders_kodu
List2.AddItem tbl!ders_adi & " - " & tbl3!ad_soyad
List5.AddItem tbl3!sicil_no
tbl2.MoveNext
Loop
If tbl4!ogretim = -1 Then
For s = 1 To 5
For i = 7 To 12
MSFlexGrid1.TextMatrix(i, s) = "X"
MSFlexGrid2.TextMatrix(i, s) = "X"
MSFlexGrid3.TextMatrix(i, s) = "X"
Next
Next
Else
For s = 1 To 5
For i = 1 To 6
MSFlexGrid1.TextMatrix(i, s) = "X"
MSFlexGrid2.TextMatrix(i, s) = "X"
MSFlexGrid3.TextMatrix(i, s) = "X"
Next
Next
End If
veri_tab.Close
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
MSFlexGrid2.TextMatrix(0, 1) = "Pazartesi"
MSFlexGrid2.TextMatrix(0, 2) = "Salý"
MSFlexGrid2.TextMatrix(0, 3) = "Çarþamba"
MSFlexGrid2.TextMatrix(0, 4) = "Perþembe"
MSFlexGrid2.TextMatrix(0, 5) = "Cuma"
MSFlexGrid3.TextMatrix(0, 1) = "Pazartesi"
MSFlexGrid3.TextMatrix(0, 2) = "Salý"
MSFlexGrid3.TextMatrix(0, 3) = "Çarþamba"
MSFlexGrid3.TextMatrix(0, 4) = "Perþembe"
MSFlexGrid3.TextMatrix(0, 5) = "Cuma"
Do While Not tbl.EOF
MSFlexGrid1.AddItem tbl!ders_saati
MSFlexGrid2.AddItem tbl!ders_saati
MSFlexGrid3.AddItem tbl!ders_saati
tbl.MoveNext
Loop

veri_tab.Close
End Sub

Private Sub MSFlexGrid1_dblClick()
On Local Error GoTo hata
If MSFlexGrid1.MouseRow = 0 Or MSFlexGrid1.MouseCol = 0 Then GoTo bitis
If MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) = "" Then GoTo bitis
If MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) = "X" Then GoTo bitis
Call vt_baglan
Call tbl_baglan("select * from sinif where sinif_kod='" & List4.List(List1.ListIndex) & "'")
''''''''''''
If tbl!ogretim = -1 Then
MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) = InputBox("Dersin Yapýlacaðý Salonu Giriniz")

If MSFlexGrid1.MouseRow <> 6 Then
If MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow + 1, MSFlexGrid1.MouseCol) = MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) Then
MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow + 1, MSFlexGrid1.MouseCol) = MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol)
End If
End If

If MSFlexGrid1.MouseRow <> 1 Then
If MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow - 1, MSFlexGrid1.MouseCol) = MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) Then
MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow - 1, MSFlexGrid1.MouseCol) = MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol)
End If
End If


Else
MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) = InputBox("Dersin Yapýlacaðý Salonu Giriniz")
'''ÇAKIÞMA
'MSFlexGrid.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol)
Call tbl2_baglan("select * from ders_prog where sinif_kod<>'" & List4.List(List1.ListIndex) & "'")

t1 = MSFlexGrid1.MouseRow
t2 = MSFlexGrid1.MouseCol
If t2 = 1 Then gun = "Pazartesi"
If t2 = 2 Then gun = "Salý"
If t2 = 3 Then gun = "Çarþamba"
If t2 = 4 Then gun = "Perþembe"
If t2 = 5 Then gun = "Cuma"

Do While Not tbl2.EOF
If MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) = salon Then
If tbl2!saat_bol = 20 Then
If t1 = tbl2!saat And tbl2!gun = gun Then
MsgBox gun & " günü " & t1 & ".saat salon " & MSFlexGrid3.TextMatrix(t1, t2) & " doludur!"
msfexgrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) = ""
GoTo bitis
End If
Else
If t1 = Val((Left(tbl2!saat, InStr(1, tbl2!saat, ",") - 1))) And gun = Left(tbl2!gun, InStr(1, tbl2!gun, ",") - 1) Then GoTo bitis
If t1 = Val((Right(tbl2!saat, Len(tbl2!saat) - InStr(1, tbl2!saat, ",")))) And gun = Right(tbl2!gun, Len(tbl2!gun) - InStr(1, tbl2!gun, ",")) Then GoTo bitis
End If
End If
tbl2.MoveNext
Loop
'''''''''''''''''''''''

If MSFlexGrid1.MouseRow <> 12 Then
If MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow + 1, MSFlexGrid1.MouseCol) = MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) Then
MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow + 1, MSFlexGrid1.MouseCol) = MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol)
End If
End If

If MSFlexGrid1.MouseRow <> 7 Then
If MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow - 1, MSFlexGrid1.MouseCol) = MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol) Then
MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow - 1, MSFlexGrid1.MouseCol) = MSFlexGrid3.TextMatrix(MSFlexGrid1.MouseRow, MSFlexGrid1.MouseCol)
End If
End If


End If
veri_tab.Close
Exit Sub
hata:
If Err = 381 Then GoTo bitis
bitis:
End Sub

Private Sub MsFlexgrid1_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MSFlexGrid1.MousePointer = 99
If MSFlexGrid1.MouseCol = 0 Or MSFlexGrid1.MouseRow = 0 Then c3 = 1

If c1 = 1 Or c2 = 1 Then c3 = 1
bas_flex_x = MSFlexGrid1.Row
bas_flex_y = MSFlexGrid1.Col
End If
End Sub

Private Sub MsFlexGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MSFlexGrid1.MousePointer = 0
hedef_flex_x = MSFlexGrid1.RowSel
hedef_flex_y = MSFlexGrid1.ColSel

If MSFlexGrid1.MouseCol = 0 Or MSFlexGrid1.MouseRow = 0 Or c3 = 1 Then
c3 = 0
GoTo bitis
End If
If c1 = 1 Or c2 = 1 Then GoTo bitis
If bas_flex_x = hedef_flex_x And bas_flex_y = hedef_flex_y Then GoTo bitis

Call degisim
End If
bitis:
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
k = 0
islem1 = 0
Do While k < MSFlexGrid1.Cols
islem1 = islem1 + MSFlexGrid1.ColWidth(k)
k = k + 1
Loop
k = 0
islem2 = 0
Do While k < MSFlexGrid1.Rows
islem2 = islem2 + MSFlexGrid1.RowHeight(k)
k = k + 1
Loop
If (X < MSFlexGrid1.Width) And (X > islem1) Then
c1 = 1
Else:
c1 = 0
End If
If (Y < MSFlexGrid1.Height) And (Y > islem2) Then
c2 = 1
Else:
c2 = 0
End If
End Sub

Sub degisim()
Dim degis As Variant

degis = MSFlexGrid1.TextMatrix(hedef_flex_x, hedef_flex_y)
MSFlexGrid1.TextMatrix(hedef_flex_x, hedef_flex_y) = MSFlexGrid1.TextMatrix(bas_flex_x, bas_flex_y)
MSFlexGrid1.TextMatrix(bas_flex_x, bas_flex_y) = degis

degis = MSFlexGrid2.TextMatrix(hedef_flex_x, hedef_flex_y)
MSFlexGrid2.TextMatrix(hedef_flex_x, hedef_flex_y) = MSFlexGrid2.TextMatrix(bas_flex_x, bas_flex_y)
MSFlexGrid2.TextMatrix(bas_flex_x, bas_flex_y) = degis

degis = MSFlexGrid3.TextMatrix(hedef_flex_x, hedef_flex_y)
MSFlexGrid3.TextMatrix(hedef_flex_x, hedef_flex_y) = MSFlexGrid3.TextMatrix(bas_flex_x, bas_flex_y)
MSFlexGrid3.TextMatrix(bas_flex_x, bas_flex_y) = degis

End Sub

Sub salon_kaydet()
Call vt_baglan
Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "'")

For s = 1 To 5
For i = 1 To 6
If MSFlexGrid3.TextMatrix(i, s) <> "" Then
Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "' and ders_kodu='" & MSFlexGrid2.TextMatrix(i, s) & "'")
tbl.Edit
tbl!salon = MSFlexGrid3.TextMatrix(i, s)
tbl.Update
End If

Next
Next

End Sub

Sub salon_kaydet2()
Call vt_baglan
Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "'")

For s = 1 To 5
For i = 7 To 12
If MSFlexGrid3.TextMatrix(i, s) <> "" Then
Call tbl_baglan("select * from ders_prog where sinif_kod='" & List4.List(List1.ListIndex) & "' and ders_kodu='" & MSFlexGrid2.TextMatrix(i, s) & "'")
tbl.Edit
tbl!salon = MSFlexGrid3.TextMatrix(i, s)
tbl.Update
End If

Next
Next
End Sub
