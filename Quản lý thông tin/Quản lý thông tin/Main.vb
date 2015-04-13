Imports System.Data.SqlClient 'Gọi thư viện
Imports System.Data.DataTable 'Gọi thư viện
Imports System.Data.DataSet
'Võ Anh Kiệt
Public Class Frmmain
    'Tạo các biến sử dụng cho toàn form
    Dim Sanpham As New DataTable ' Biến của bảng sản phẩm
    Dim nhanvien As New DataTable ' Biếm của bảng nhân viên
    Dim KhachHang As New DataTable ' Biến của bảng khách hàng
    Dim hoadon As New DataTable ' Biến của bảng hóa đơn
    Dim mua As New DataTable
    'Biến dẫn đến cơ sỡ dữ liệu đám mây
    Dim conmectstr As String = "workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 "
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    Public Sub Mua_load() ' Sự kiên load form sản phẩm
        Dim KetNoiMua As New SqlConnection(conmectstr)
        Dim sqlSanPham As New SqlDataAdapter("select MaSanPham as  'Mã Sản Phẩm',Ten as 'Tên Sản Phẩm',Loai as 'Loại Sản Phẩm',Don as 'Đơn Giá',So as 'Số Lượng',Nha as 'Nhà Sản Xuât',Ngay 'Ngày Sản Xuất',Ghi as 'Ghi Chú' from SanPham", KetNoiMua)
        Try
            KetNoiMua.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlSanPham.Fill(mua)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvMua.DataSource = mua
        KetNoiMua.Close()
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    Public Sub hoadon_load() ' Sự kiên load form hóa đơn
        Dim KetNoiHD As New SqlConnection(conmectstr)
        Dim sqlHoadon As New SqlDataAdapter("select MaHD as 'Mã Hóa Đơn',Gia as 'Đơn Giá',So as 'Số Lượng',DC as 'Địa Chỉ',Tong as 'Thành Tiền',Ngay as 'Ngày Mua' from hoadon", KetNoiHD)
        Try
            KetNoiHD.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlHoadon.Fill(hoadon)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvhoadon.DataSource = hoadon
        KetNoiHD.Close()
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    Public Sub Khachhang_load() ' Sự kiên load form hóa đơn
        Dim KetNoiKH As New SqlConnection(conmectstr)
        Dim sqlKhachhang As New SqlDataAdapter("select MaKH as 'Mã Khách Hàng',TenKH as 'Tên Khách Hàng',Loai as 'Loại Khách Hàng',DiaChi as 'Địa Chỉ',Email,Dienthoai as 'Điện Thoại',Ngay as 'Ngày Gia Nhập',Ghichu as 'Ghi Chú' from khachhang", KetNoiKH)
        Try
            KetNoiKH.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlKhachhang.Fill(KhachHang)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvKH.DataSource = KhachHang
        KetNoiKH.Close()
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    Public Sub nhanvien_load() ' Sự kiên load form hóa đơn
        Dim KetNoiNV As New SqlConnection(conmectstr)
        Dim sqlnhanvien As New SqlDataAdapter("select Manhanvien as 'Mã Nhân Viên',HoNhanvien as 'Họ Nhân Viên',TenNhanvien as'Tên Nhân viên',Taikhoan as 'Tài Khoản',Matkhau as 'Mật Khẩu',ChucVu as 'Chức Vụ',NgayBD as 'Ngày Bắt Đầu',DiaChi as 'Địa Chỉ',Email as 'Email',SDT as'SĐT',Nganhang as 'Ngân Hàng',STK as 'Số Tài Khoản' from nhanvien", KetNoiNV)
        Try
            KetNoiNV.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlnhanvien.Fill(nhanvien)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvNhanVien.DataSource = nhanvien
        KetNoiNV.Close()
    End Sub
    Private Sub NhânViênToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'frmNhanVien.Show() 'Sự kiện form nhân viên được show ra
    End Sub
    'Sự kiện sao khi nhấp vào nút thoát
    Private Sub ThoátToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThoátToolStripMenuItem.Click
        Me.Close()
        Login.Close()
    End Sub
    'Sự kiện sao khi nhấp vào nút đăng xuất
    Private Sub DDaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DDaToolStripMenuItem.Click
        Me.Close()
        Login.Show()
    End Sub

    Private Sub Frmmain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '--------------------------------------------------------------
        Dim KetNoiNV As New SqlConnection(conmectstr)
        Dim sqlnhanvien As New SqlDataAdapter("select Manhanvien as 'Mã Nhân Viên',HoNhanvien as 'Họ Nhân Viên',TenNhanvien as'Tên Nhân viên',Taikhoan as 'Tài Khoản',Matkhau as 'Mật Khẩu',ChucVu as 'Chức Vụ',NgayBD as 'Ngày Bắt Đầu',DiaChi as 'Địa Chỉ',Email as 'Email',SDT as'SĐT',Nganhang as 'Ngân Hàng',STK as 'Số Tài Khoản' from nhanvien", KetNoiNV)
        Try
            KetNoiNV.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlnhanvien.Fill(nhanvien)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvNhanVien.DataSource = nhanvien
        KetNoiNV.Close()
        ''--------------------------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------------------
        ' Sự kiện kiad dữ liệu kên combo
        Dim a As New Combox
        cboMaKHBan.DataSource = a.Load("MaKH").Tables(0)
        cboMaKHBan.DisplayMember = "MaKH"
        ComboBox1.DataSource = a.Load3("MaNhanvien").Tables(0)
        ComboBox1.DisplayMember = "MaNhanvien"
        ''--------------------------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------------------
        Dim KetNoiHD As New SqlConnection(conmectstr)
        Dim sqlHoadon As New SqlDataAdapter("select MaHD as 'Mã Hóa Đơn',Gia as 'Đơn Giá',So as 'Số Lượng',DC as 'Địa Chỉ',Tong as 'Thành Tiền',Ngay as 'Ngày Mua' from hoadon", KetNoiHD)
        Try
            KetNoiHD.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlHoadon.Fill(hoadon)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvhoadon.DataSource = hoadon
        KetNoiHD.Close()
        '--------------------------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        'Sự Kiện Gọi Các cột của trong Bảng Sản phẩm
        Dim KetNoiMua As New SqlConnection(conmectstr)
        Dim sqlMua As New SqlDataAdapter("select MaSanPham as  'Mã Sản Phẩm',Ten as 'Tên Sản Phẩm',Loai as 'Loại Sản Phẩm',Don as 'Đơn Giá',So as 'Số Lượng',Nha as 'Nhà Sản Xuât',Ngay 'Ngày Sản Xuất',Ghi as 'Ghi Chú' from SanPham", KetNoiMua)
        Try
            KetNoiMua.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlMua.Fill(mua)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvMua.DataSource = mua
        dgvBanSP.DataSource = mua
        KetNoiMua.Close()
        '--------------------------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        'Sự Kiện Gọi Các cột của trong Bảng Sản phẩm

        '--------------------------------------------------------------------------------------------------------
        Dim KetNoiKH As New SqlConnection(conmectstr)
        Dim sqlKhachhang As New SqlDataAdapter("select MaKH as 'Mã Khách Hàng',TenKH as 'Tên Khách Hàng',Loai as 'Loại Khách Hàng',DiaChi as 'Địa Chỉ',Email,Dienthoai as 'Điện Thoại',Ngay as 'Ngày Gia Nhập',Ghichu as 'Ghi Chú' from khachhang", KetNoiKH)
        Try
            KetNoiKH.Open()
            'Đổ dữ liệu trên Table vào Datatable trên máy
            sqlKhachhang.Fill(KhachHang)
        Catch ex As Exception

        End Try
        'Hiển thị dữ liệu Từ Datatable ra giao diện thông qua Datagridview
        dgvKH.DataSource = KhachHang
        KetNoiKH.Close()
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    Private Sub dgvBanSP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBanSP.CellContentClick
        Dim index As Integer = dgvBanSP.CurrentCell.RowIndex ' Tạo biến
        TextBox1.Text = dgvBanSP.Item(0, index).Value
        txtDGBan.Text = dgvBanSP.Item(3, index).Value
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    'Sự kiện sao khi nhấp vào 1 đối tượng tại Datagitview Mua
    Private Sub dgvMua_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvMua.CellContentClick
        Dim index As Integer = dgvMua.CurrentCell.RowIndex
        txtMHmua.Text = dgvMua.Item(0, index).Value 'Đưa lên textbox 1
        txtTSPMua.Text = dgvMua.Item(1, index).Value 'Đưa lên textbox 2
        cboLHMua.Text = dgvMua.Item(2, index).Value 'Đưa lên textbox 3
        txtSLmua.Text = dgvMua.Item(3, index).Value 'Đưa lên textbox 4
        txtDGMua.Text = dgvMua.Item(4, index).Value 'Đưa lên textbox 5
        txtNSXMua.Text = dgvMua.Item(5, index).Value 'Đưa lên textbox 6
        txtNgSXMua.Text = dgvMua.Item(6, index).Value 'Đưa lên textbox 7
        txtGCMua.Text = dgvMua.Item(7, index).Value 'Đưa lên textbox 7
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    'Sự kiện sao khi nhấp vào 1 đối tượng tại Datagitview KhachHanng
    Private Sub dgvKH_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvKH.CellContentClick
        Dim index As Integer = dgvKH.CurrentCell.RowIndex
        txtMaKHKH.Text = dgvKH.Item(0, index).Value ' Đưa lên textbox 1
        txtTenKHKH.Text = dgvKH.Item(1, index).Value ' Đưa lên textbox 2
        txtLKHKH.Text = dgvKH.Item(2, index).Value ' Đưa lên textbox 3
        txtDCKH.Text = dgvKH.Item(3, index).Value ' Đưa lên textbox 4
        txtEmailKH.Text = dgvKH.Item(4, index).Value ' Đưa lên textbox 5
        txtDTKH.Text = dgvKH.Item(5, index).Value ' Đưa lên textbox 6
        dtpNgayKHKH.Text = dgvKH.Item(6, index).Value ' Đưa lên textbox 7
        txtGhiChuKH.Text = dgvKH.Item(7, index).Value ' Đưa lên textbox 8

    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    'Sự kiện nhấp vào nút Thêm hàng ở bản mua
    Private Sub btnThemMua_Click(sender As Object, e As EventArgs) Handles btnThemMua.Click
        Dim query As String = String.Empty
        query &= "INSERT INTO SanPham (MaSanPham,Ten,Loai,Don,So,Nha,Ngay,Ghi)"
        query &= "Values(@MaSp,@Ten,@Loai,@Don,@Gia,@Nha,@Ngay,@Ghi)"
        Using conn As New SqlConnection("workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 ")
            Using comm As New SqlCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MaSP", txtMHmua.Text)
                    .Parameters.AddWithValue("@Ten", txtTSPMua.Text)
                    .Parameters.AddWithValue("@Loai", cboLHMua.Text)
                    .Parameters.AddWithValue("@Don", txtDGMua.Text)
                    .Parameters.AddWithValue("@Gia", txtSLmua.Text)
                    .Parameters.AddWithValue("@Nha", txtNSXMua.Text)
                    .Parameters.AddWithValue("@Ngay", txtNgSXMua.Text)
                    .Parameters.AddWithValue("@Ghi", txtGCMua.Text)
                End With
                Try
                    conn.Open()
                    comm.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error Message")
                End Try
                mua.Clear()
                dgvMua.DataSource = mua
                dgvMua.DataSource = Nothing
                Mua_load()
            End Using
        End Using
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    'Sự kiện khi nhấp vào nút xóa ở bản Mua
    Private Sub btnXoaMua_Click(sender As Object, e As EventArgs) Handles btnXoaMua.Click
        Dim KetNoi As New SqlConnection(conmectstr)
        KetNoi.Open()
        Dim stradd As String = "Delete from sanpham Where MaSanPham = @MaSP"
        Dim com As New SqlCommand(stradd, KetNoi)
        Try
            com.Parameters.AddWithValue("@MaSp", txtMHmua.Text)
            com.ExecuteNonQuery()
            KetNoi.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "Error Message")
        End Try
        mua.Clear()
        dgvMua.DataSource = mua
        dgvMua.DataSource = Nothing
        Mua_load()
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    'Sự kiện nhấn vào nút sưa ở bản mua
    Private Sub btnSuaMua_Click(sender As Object, e As EventArgs) Handles btnSuaMua.Click
        Dim KetNoi As New SqlConnection(conmectstr)
        KetNoi.Open()
        Dim stradd As String = "Update sanpham set Ten =@Ten, Loai =@Loai, Don =@Don ,So =@So,Nha =@Nha,Ngay =@Ngay , Ghi =@Ghi where MaSanPham =@MaSP "
        Dim Com As New SqlCommand(stradd, KetNoi)
        Try
            Com.Parameters.AddWithValue("@MaSP", txtMHmua.Text)
            Com.Parameters.AddWithValue("@Ten", txtTSPMua.Text)
            Com.Parameters.AddWithValue("@Loai", cboLHMua.Text)
            Com.Parameters.AddWithValue("@Don", txtDGMua.Text)
            Com.Parameters.AddWithValue("@So", txtSLmua.Text)
            Com.Parameters.AddWithValue("@Nha", txtNSXMua.Text)
            Com.Parameters.AddWithValue("@Ngay", txtNgSXMua.Text)
            Com.Parameters.AddWithValue("@Ghi", txtGCMua.Text)
            Com.ExecuteNonQuery()
            KetNoi.Close()
        Catch ex As Exception
            MessageBox.Show("Error")
        End Try
        mua.Clear()
        dgvMua.DataSource = mua
        dgvMua.DataSource = Nothing
        Mua_load()
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------

    Private Sub btnLammoimua_Click(sender As Object, e As EventArgs) Handles btnLammoimua.Click
        txtMHmua.Text = ""
        txtTSPMua.Text = ""
        cboLHMua.Text = ""
        txtSLmua.Text = ""
        txtDGMua.Text = ""
        txtNSXMua.Text = ""
        txtNgSXMua.Text = ""
        txtGCMua.Text = ""
    End Sub
    'Sự kiện khi nhấn vào nút thêm thành viên ở bảng khách hàng
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim query As String = String.Empty
        query &= "INSERT INTO khachhang (MaKH,TenKH,Loai,Diachi,email,dienthoai,Ngay,Ghichu)"
        query &= "Values(@MaKH,@TenKH,@Loai,@DC,@Email,@DT,@Ngay,@Ghi)"
        Using conn As New SqlConnection("workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 ")
            Using comm As New SqlCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MaKH", txtMaKHKH.Text)
                    .Parameters.AddWithValue("@TenKH", txtTenKHKH.Text)
                    .Parameters.AddWithValue("@Loai", txtLKHKH.Text)
                    .Parameters.AddWithValue("@DC", txtDCKH.Text)
                    .Parameters.AddWithValue("@Email", txtEmailKH.Text)
                    .Parameters.AddWithValue("@DT", txtDTKH.Text)
                    .Parameters.AddWithValue("@Ngay", dtpNgayKHKH.Text)
                    .Parameters.AddWithValue("@Ghi", txtGhiChuKH.Text)
                End With
                Try
                    conn.Open()
                    comm.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error Message")
                End Try
                KhachHang.Clear()
                dgvKH.DataSource = KhachHang
                dgvKH.DataSource = Nothing
                Khachhang_load()
            End Using
        End Using
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    'Sự kiện nhấp vào nút sửa bảng khách hàng
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim KetNoi As New SqlConnection(conmectstr)
        KetNoi.Open()
        Dim stradd As String = "Update Khachhang set TenKH =@Ten,Loai =@Loai, DiaCHi =@DC ,Email =@Email,Dienthoai =@DT ,Ngay =@Ngay ,Ghichu =@Ghi where MaKH =@MaKH "
        Dim Com As New SqlCommand(stradd, KetNoi)
        Try
            Com.Parameters.AddWithValue("@MaKH", txtMaKHKH.Text)
            Com.Parameters.AddWithValue("@Ten", txtTenKHKH.Text)
            Com.Parameters.AddWithValue("@Loai", txtLKHKH.Text)
            Com.Parameters.AddWithValue("@DC", txtDCKH.Text)
            Com.Parameters.AddWithValue("@Email", txtEmailKH.Text)
            Com.Parameters.AddWithValue("@DT", txtDTKH.Text)
            Com.Parameters.AddWithValue("@Ngay", dtpNgayKHKH.Text)
            Com.Parameters.AddWithValue("@Ghi", txtGhiChuKH.Text)
            Com.ExecuteNonQuery()
            KetNoi.Close()
        Catch ex As Exception
            MessageBox.Show("Error")
        End Try
        KhachHang.Clear()
        dgvKH.DataSource = KhachHang
        dgvKH.DataSource = Nothing
        Khachhang_load()
    End Sub
    'Sự kiện nhấp vào nút xóa
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim KetNoi As New SqlConnection(conmectstr)
        KetNoi.Open()
        Dim stradd As String = "Delete from khachhang Where MaKH = @MaKH"
        Dim com As New SqlCommand(stradd, KetNoi)
        Try
            com.Parameters.AddWithValue("MaKH", txtMaKHKH.Text)
            com.ExecuteNonQuery()
            KetNoi.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "Error Message")
        End Try
        KhachHang.Clear()
        dgvKH.DataSource = KhachHang
        dgvKH.DataSource = Nothing
        Khachhang_load()
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    'Sự kiện sao nhi nhấm vào nút làm mới Khách Hàng
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        txtMaKHKH.Text = ""
        txtTenKHKH.Text = ""
        txtLKHKH.Text = ""
        txtDCKH.Text = ""
        txtEmailKH.Text = ""
        txtDTKH.Text = ""
        dtpNgayKHKH.Text = ""
        txtGhiChuKH.Text = ""
    End Sub
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------
    ' Sự kiện nhấp vào nút thêm ở bản Bán
    Private Sub btnThemBan_Click(sender As Object, e As EventArgs) Handles btnThemBan.Click
        Dim query As String = String.Empty
        query &= "INSERT INTO HoaDon (MaHD,Gia,So,DC,tong,ngay,nhanvien_manhanvien,sanpham_masanpham,khachhang_makh)"
        query &= "Values(@MaHD,@Gia,@So,@DC,@TT,@ngay,@NV,@SP,@kh)"
        Using conn As New SqlConnection("workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 ")
            Using comm As New SqlCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@MaHD", txtMaHDBan.Text)
                    .Parameters.AddWithValue("@Gia", txtDGBan.Text)
                    .Parameters.AddWithValue("@So", txtSLBan.Text)
                    .Parameters.AddWithValue("@DC", txtDCban.Text)
                    .Parameters.AddWithValue("@TT", txtTTBan.Text)
                    .Parameters.AddWithValue("@ngay", dtbban.Text)
                    .Parameters.AddWithValue("@NV", ComboBox1.Text)
                    .Parameters.AddWithValue("@SP", TextBox1.Text)
                    .Parameters.AddWithValue("@kh", cboMaKHBan.Text)
                End With
                Try
                    conn.Open()
                    comm.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error Message")
                End Try
                hoadon.Clear()
                dgvhoadon.DataSource = hoadon
                dgvhoadon.DataSource = Nothing
                hoadon_load()
            End Using
        End Using
    End Sub

    Private Sub cboMaSPBan_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub dgvNhanVien_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvNhanVien.CellContentClick
        Dim index As Integer = dgvNhanVien.CurrentCell.RowIndex
        TextBox22.Text = dgvNhanVien.Item(0, index).Value
        TextBox21.Text = dgvNhanVien.Item(1, index).Value
        TextBox20.Text = dgvNhanVien.Item(2, index).Value
        TextBox19.Text = dgvNhanVien.Item(3, index).Value
        TextBox18.Text = dgvNhanVien.Item(4, index).Value
        TextBox17.Text = dgvNhanVien.Item(5, index).Value
        TextBox16.Text = dgvNhanVien.Item(6, index).Value
        TextBox15.Text = dgvNhanVien.Item(7, index).Value
        TextBox14.Text = dgvNhanVien.Item(8, index).Value
        TextBox13.Text = dgvNhanVien.Item(9, index).Value
        TextBox12.Text = dgvNhanVien.Item(10, index).Value
        TextBox11.Text = dgvNhanVien.Item(11, index).Value
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim query As String = String.Empty
        query &= "INSERT INTO nhanvien (Manhanvien,HoNhanvien,TenNhanvien,Taikhoan,matkhau,Chucvu,NgayBD,DiaChi,Email,SDT,Nganhang,STK)"
        query &= "Values(@Manv,@Ho,@Ten,@Tai,@mat,@Chuc,@Ngay,@DC,@Em,@SDT,@Ngan,@Stk)"
        Using conn As New SqlConnection("workstation id=anhkietcm4.mssql.somee.com;packet size=4096;user id=anhkietcm20;pwd=Abc01091989491;data source=anhkietcm4.mssql.somee.com;persist security info=False;initial catalog=anhkietcm4 ")
            Using comm As New SqlCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.AddWithValue("@Manv", TextBox22.Text)
                    .Parameters.AddWithValue("@Ho", TextBox21.Text)
                    .Parameters.AddWithValue("@Ten", TextBox20.Text)
                    .Parameters.AddWithValue("@Tai", TextBox19.Text)
                    .Parameters.AddWithValue("@mat", TextBox18.Text)
                    .Parameters.AddWithValue("@Chuc", TextBox17.Text)
                    .Parameters.AddWithValue("@Ngay", TextBox16.Text)
                    .Parameters.AddWithValue("@DC", TextBox15.Text)
                    .Parameters.AddWithValue("@Em", TextBox14.Text)
                    .Parameters.AddWithValue("@SDT", TextBox14.Text)
                    .Parameters.AddWithValue("@Ngan", TextBox14.Text)
                    .Parameters.AddWithValue("@Stk", TextBox14.Text)
                End With
                Try
                    conn.Open()
                    comm.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString(), "Error Message")
                End Try
                nhanvien.Clear()
                dgvNhanVien.DataSource = hoadon
                dgvNhanVien.DataSource = Nothing
                nhanvien_load()
            End Using
        End Using
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim KetNoi As New SqlConnection(conmectstr)
        KetNoi.Open()
        Dim stradd As String = "Delete from nhanvien Where Manhanvien = @MaNV"
        Dim com As New SqlCommand(stradd, KetNoi)
        Try
            com.Parameters.AddWithValue("@Manv", TextBox22.Text)
            com.ExecuteNonQuery()
            KetNoi.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "Error Message")
        End Try
        nhanvien.Clear()
        dgvNhanVien.DataSource = nhanvien
        dgvNhanVien.DataSource = Nothing
        nhanvien_load()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox22.Text = ""
        TextBox21.Text = ""
        TextBox20.Text = ""
        TextBox19.Text = ""
        TextBox18.Text = ""
        TextBox17.Text = ""
        TextBox16.Text = ""
        TextBox15.Text = ""
        TextBox14.Text = ""
        TextBox13.Text = ""
        TextBox12.Text = ""
        TextBox11.Text = ""
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim KetNoi As New SqlConnection(conmectstr)
        KetNoi.Open()
        Dim stradd As String = "Update nhanvien set HoNhanvien =@Ho,TenNhanvien =@Ten,TaiKhoan =@Tai,matkhau =@mat,Chucvu =@Chuc,ngaybd =@Ngay,Diachi =@DC,Email =@Em,SDT =@SDT,nganhang =@Ngan,STK =@Stk where Manhanvien =@Manv"
        Dim Com As New SqlCommand(stradd, KetNoi)
        Try
            Com.Parameters.AddWithValue("@Manv", TextBox22.Text)
            Com.Parameters.AddWithValue("@Ho", TextBox21.Text)
            Com.Parameters.AddWithValue("@Ten", TextBox20.Text)
            Com.Parameters.AddWithValue("@Tai", TextBox19.Text)
            Com.Parameters.AddWithValue("@mat", TextBox18.Text)
            Com.Parameters.AddWithValue("@Chuc", TextBox17.Text)
            Com.Parameters.AddWithValue("@Ngay", TextBox16.Text)
            Com.Parameters.AddWithValue("@DC", TextBox15.Text)
            Com.Parameters.AddWithValue("@Em", TextBox14.Text)
            Com.Parameters.AddWithValue("@SDT", TextBox14.Text)
            Com.Parameters.AddWithValue("@Ngan", TextBox14.Text)
            Com.Parameters.AddWithValue("@Stk", TextBox14.Text)
            Com.ExecuteNonQuery()
            KetNoi.Close()
        Catch ex As Exception
            MessageBox.Show("Error")
        End Try
        nhanvien.Clear()
        dgvNhanVien.DataSource = nhanvien
        dgvNhanVien.DataSource = Nothing
        nhanvien_load()
    End Sub

    Private Sub txtSLBan_TextChanged(sender As Object, e As EventArgs) Handles txtSLBan.TextChanged
        txtTTBan.Text = txtSLBan.Text * txtDGBan.Text
    End Sub

    Private Sub btnSuaBan_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub dgvhoadon_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvhoadon.CellContentClick
        Dim index As Integer = dgvhoadon.CurrentCell.RowIndex
        TextBox2.Text = dgvhoadon.Item(0, index).Value
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim KetNoi As New SqlConnection(conmectstr)
        KetNoi.Open()
        Dim stradd As String = "Delete from hoadon Where MaHD = @MaHD"
        Dim com As New SqlCommand(stradd, KetNoi)
        Try
            com.Parameters.AddWithValue("@MaHD", TextBox2.Text)
            com.ExecuteNonQuery()
            KetNoi.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "Error Message")
        End Try
        hoadon.Clear()
        dgvhoadon.DataSource = hoadon
        dgvhoadon.DataSource = Nothing
        hoadon_load()
    End Sub

    Private Sub Frmmain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Login.Close()
    End Sub
End Class
