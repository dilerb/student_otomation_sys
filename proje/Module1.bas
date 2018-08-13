Attribute VB_Name = "Module1"
Global veri_tab As Database
Global tbl As Recordset
Global tbl2 As Recordset
Global tbl3 As Recordset
Global tbl4 As Recordset
Global tbl5 As Recordset
Global tbl6 As Recordset
Global tbl7 As Recordset

Sub vt_baglan()
Set veri_tab = Workspaces(0).OpenDatabase(App.Path & "\dersmatik.mdb", False, False, ";pwd=123")
End Sub

Sub tbl_baglan(sql As String)
Set tbl = veri_tab.OpenRecordset(sql)
End Sub

Sub tbl2_baglan(sql As String)
Set tbl2 = veri_tab.OpenRecordset(sql)
End Sub

Sub tbl3_baglan(sql As String)
Set tbl3 = veri_tab.OpenRecordset(sql)
End Sub

Sub tbl4_baglan(sql As String)
Set tbl4 = veri_tab.OpenRecordset(sql)
End Sub

Sub tbl5_baglan(sql As String)
Set tbl5 = veri_tab.OpenRecordset(sql)
End Sub

Sub tbl6_baglan(sql As String)
Set tbl6 = veri_tab.OpenRecordset(sql)
End Sub

Sub tbl7_baglan(sql As String)
Set tbl7 = veri_tab.OpenRecordset(sql)
End Sub
