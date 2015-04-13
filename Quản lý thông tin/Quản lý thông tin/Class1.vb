Imports System.Data.SqlClient
Imports System.Data.DataTable
Public Class Combox

    Function Load(p1 As String) As Object
        Dim S As String = "workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 "
        Dim conn As New SqlConnection(S)
        Dim da As New SqlDataAdapter("select * from khachhang", conn)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds)
        conn.Close()
        Return ds
    End Function
    Function Load2(p1 As String) As Object
        Dim S As String = "workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 "
        Dim conn As New SqlConnection(S)
        Dim da As New SqlDataAdapter("select * from sanpham", conn)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds)
        conn.Close()
        Return ds
    End Function
    Function Load3(p1 As String) As Object
        Dim S As String = "workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 "
        Dim conn As New SqlConnection(S)
        Dim da As New SqlDataAdapter("select * from nhanvien", conn)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds)
        conn.Close()
        Return ds
    End Function


End Class
