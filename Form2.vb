Imports System.Data.OleDb
Imports System.Globalization
Public Class Form2
    Dim koneksi As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\azriel hasan\Desktop\New folder\Acces Data\dbMahasiswa.accdb")
    Dim dr As OleDbDataReader

    Sub loadingDataGridView()
        Try
            DataGridView1.Rows.Clear()
            koneksi.Open()
            Dim cmd As New OleDbCommand("Select * from Mahasiswa ", koneksi)
            dr = cmd.ExecuteReader
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("NPM"), dr.Item("Nama"), dr.Item("Alamat"), dr.Item("JenisKelamin"), dr.Item("TanggalLahir"),
                                    dr.Item("Jurusan"), dr.Item("Agama"), dr.Item("NoTelp"))
            End While
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        koneksi.Close()
    End Sub
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Isi ComboBox Jenis Kelamin
        cmbJK.Items.Add("Laki Laki")
        cmbJK.Items.Add("Perempuan")

        ' Isi ComboBox Agama
        cmbAgama.Items.Add("Islam")
        cmbAgama.Items.Add("Kristen")
        cmbAgama.Items.Add("Katolik")
        cmbAgama.Items.Add("Hindu")
        cmbAgama.Items.Add("Buddha")
        cmbAgama.Items.Add("Konghucu")

        ' Isi ComboBox Jurusan
        cmbJurusan.Items.Add("Manajemen")
        cmbJurusan.Items.Add("Akutansi")
        cmbJurusan.Items.Add("Ekonomi")

        loadingDataGridView() ' prosedur untuk load data dari database ke DataGridView
    End Sub

    'tampilkan data



    Private Sub btnSimpan_Click(sender As Object, e As EventArgs) Handles btnSimpan.Click
        DataGridView1.Rows.Add(txtNpm.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text, dtpTanggalLahir.Text, cmbJurusan.Text, cmbAgama.Text, txtTelp.Text)
        Try
            koneksi.Open()
            Dim cmd As New OleDbCommand("INSERT INTO Mahasiswa (NPM, Nama, Alamat, JenisKelamin, TanggalLahir, Jurusan, Agama, NoTelp) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", koneksi)
            cmd.Parameters.AddWithValue("?", txtNpm.Text)
            cmd.Parameters.AddWithValue("?", txtNama.Text)
            cmd.Parameters.AddWithValue("?", txtAlamat.Text)
            cmd.Parameters.AddWithValue("?", cmbJK.Text)
            Dim tglLahir As String = dtpTanggalLahir.Value.ToString("yyyy-MM-dd")
            cmd.Parameters.AddWithValue("?", tglLahir)
            cmd.Parameters.AddWithValue("?", cmbJurusan.Text)
            cmd.Parameters.AddWithValue("?", cmbAgama.Text)
            cmd.Parameters.AddWithValue("?", txtTelp.Text)
            cmd.ExecuteNonQuery()


            MessageBox.Show("Data berhasil disimpan")
            'TampilkanData() ' Refresh DataGridView
        Catch ex As Exception
            MessageBox.Show("Gagal menyimpan data: " & ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        txtNpm.Text = ""
        txtNama.Text = ""
        txtAlamat.Text = ""
        cmbJK.Text = "" 
        dtpTanggalLahir.Text =""
        cmbJurusan.Text ="" 
        cmbAgama.Text ="" 
        txtTelp.Text =""
        Try
            koneksi.Open()
            Dim cmd As New OleDbCommand("UPDATE Mahasiswa SET Nama=?, Alamat=?, JenisKelamin=?, TanggalLahir=?, Jurusan=?, Agama=?, NoTelp, WHERE NPM=?", koneksi)
            cmd.Parameters.AddWithValue("?", txtNama.Text)
            cmd.Parameters.AddWithValue("?", txtAlamat.Text)
            cmd.Parameters.AddWithValue("?", cmbJK.Text)
            cmd.Parameters.AddWithValue("?", dtpTanggalLahir.Value)
            cmd.Parameters.AddWithValue("?", cmbJurusan.Text)
            cmd.Parameters.AddWithValue("?", cmbAgama.Text)
            cmd.Parameters.AddWithValue("?", txtTelp.Text)
            cmd.Parameters.AddWithValue("?", txtNpm.Text) ' NPM sebagai primary key
            cmd.ExecuteNonQuery()


            MessageBox.Show("Data berhasil diubah")
            loadingDataGridView()
        Catch ex As Exception
            MessageBox.Show("Gagal mengubah data: " & ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click
        Try
            koneksi.Open()
            Dim cmd As New OleDb.OleDbDataAdapter("SELECT * FROM Mahasiswa WHERE NPM LIKE '%" & txtCari.Text & "%' OR Nama LIKE '%" & txtCari.Text & "%'", koneksi)
            Dim dt As New DataTable
            cmd.Fill(dt)
            DataGridView1.DataSource = dt
            koneksi.Close()
        Catch ex As Exception
            MessageBox.Show("Gagal mencari data: " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer = DataGridView1.CurrentRow.Index
        txtNpm.Text = DataGridView1.Item(0, i).Value
        txtNama.Text = DataGridView1.Item(1, i).Value
        txtAlamat.Text = DataGridView1.Item(2, i).Value
        cmbJK.Text = DataGridView1.Item(3, i).Value
        dtpTanggalLahir.Value = DataGridView1.Item(4, i).Value
        cmbJurusan.Text = DataGridView1.Item(5, i).Value
        cmbAgama.Text = DataGridView1.Item(6, i).Value
        txtTelp.Text = DataGridView1.Item(7, i).Value
    End Sub
End Class