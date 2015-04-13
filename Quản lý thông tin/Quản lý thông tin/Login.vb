Imports System.Data.SqlClient ' Câu lệnh truy cập thư viên
Public Class Login
    'Sự kiện sao khi nhấp vào button đăng nhập
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim chuoiketnoi As String = "workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 "

        Dim KetNoi As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim sqlAdapter As New SqlDataAdapter("select * from NhanVien where taikhoan='" & txtTaikhoang.Text & "' and matkhau='" & txtmatKhau.Text & "' ", KetNoi)
        Dim tb As New DataTable

        Try
            KetNoi.Open()
            sqlAdapter.Fill(tb)
            If tb.Rows.Count > 0 Then
                Frmmain.Show()
                Me.Hide()
            Else
                MessageBox.Show("Sai tài khoản hoặc mật khẩu")
            End If

        Catch ex As Exception

        End Try
    End Sub
    'Sự kiện sao khi nhấp vào nút thoát
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


End Class