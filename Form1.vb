
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Form1
    Dim strkon As String = "server=localhost;uid=root;database=laelca"
    Dim konn As New MySqlConnection(strkon)
    Dim perintah As New MySqlCommand
    Dim mda As New MySqlDataAdapter
    Dim ds As New DataSet
    Dim cek As MySqlDataReader

    Sub tidakaktif()

        txtsku.Enabled = False
        txttitle.Enabled = False
        cmbcategory.Enabled = False
        txtdeskripsi.Enabled = False
        cbs.Enabled = False
        cbm.Enabled = False
        cbl.Enabled = False
        cbxl.Enabled = False
        cbxxl.Enabled = False
        cbxxxl.Enabled = False
        txtcolour.Enabled = False
        txtmaterials.Enabled = False
        txtprice.Enabled = False
        txtstock.Enabled = False
        BtnBrowse.Enabled = False
        pbgambar.Enabled = False

        txtsku.BackColor = Color.Gray
        txttitle.BackColor = Color.Gray
        cmbcategory.BackColor = Color.Gray
        txtdeskripsi.BackColor = Color.Gray
        cballsize.BackColor = Color.Gray
        cbs.BackColor = Color.Gray
        cbm.BackColor = Color.Gray
        cbl.BackColor = Color.Gray
        cbxl.BackColor = Color.Gray
        cbxxl.BackColor = Color.Gray
        cbxxxl.BackColor = Color.Gray
        txtcolour.BackColor = Color.Gray
        txtmaterials.BackColor = Color.Gray
        txtprice.BackColor = Color.Gray
        txtstock.BackColor = Color.Gray
        BtnBrowse.BackColor = Color.Gray
        pbgambar.BackColor = Color.Gray

        cmdsimpan.Enabled = False
        cmdhapus.Enabled = False
        cmdupdate.Enabled = False
        cmdbatal.Enabled = False
    End Sub

    Sub aktif()
        txtsku.Enabled = True
        txttitle.Enabled = True
        cmbcategory.Enabled = True
        txtdeskripsi.Enabled = True
        cballsize.Enabled = True
        cbs.Enabled = True
        cbm.Enabled = True
        cbl.Enabled = True
        cbxl.Enabled = True
        cbxxl.Enabled = True
        cbxxxl.Enabled = True
        txtcolour.Enabled = True
        txtmaterials.Enabled = True
        txtprice.Enabled = True
        txtstock.Enabled = True
        BtnBrowse.Enabled = True
        pbgambar.Enabled = True

        txtsku.BackColor = Color.White
        txttitle.BackColor = Color.White
        cmbcategory.BackColor = Color.White
        txtdeskripsi.BackColor = Color.White
        cballsize.BackColor = Color.White
        cbs.BackColor = Color.White
        cbm.BackColor = Color.White
        cbl.BackColor = Color.White
        cbxl.BackColor = Color.White
        cbxxl.BackColor = Color.White
        cbxxxl.BackColor = Color.White
        txtcolour.BackColor = Color.White
        txtmaterials.BackColor = Color.White
        txtprice.BackColor = Color.White
        txtstock.BackColor = Color.White
        BtnBrowse.BackColor = Color.White
        pbgambar.BackColor = Color.White

        cmdsimpan.Enabled = True
        cmdhapus.Enabled = True
        cmdupdate.Enabled = True
        cmdbatal.Enabled = True
    End Sub

    Sub bersih()
        txtsku.Text = ""
        txttitle.Text = ""
        cmbcategory.Text = ""
        txtdeskripsi.Text = ""
        txtcolour.Text = ""
        txtmaterials.Text = ""
        txtprice.Text = ""
        txtstock.Text = ""
    End Sub

    Sub tampil()
        konn.Open()
        perintah.Connection = konn
        perintah.CommandType = CommandType.Text
        perintah.CommandText = "select * from product"
        mda.SelectCommand = perintah
        ds.Tables.Clear()
        mda.Fill(ds, "product")
        dgtampil.DataSource = ds.Tables("product")
        konn.Close()
    End Sub


    Private Sub BtnBrowse_Click(sender As Object, e As EventArgs) Handles BtnBrowse.Click
        Dim opf As New OpenFileDialog
        opf.Filter = "Choose Image (*.JPG;*.PNG;*.GIF;*.AVIF)|*.jpg;*.png;*.gif;*.avif"
        If opf.ShowDialog() = DialogResult.OK Then
            pbgambar.SizeMode = PictureBoxSizeMode.Zoom
            pbgambar.Image = Image.FromFile(opf.FileName)

            ' Simpan gambar ke MemoryStream
            Dim ms As New MemoryStream
            pbgambar.Image.Save(ms, pbgambar.Image.RawFormat)
        End If
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            konn.Open()
            Dim command As New MySqlCommand("SELECT * FROM product", konn)
            Dim adapter As New MySqlDataAdapter(command)
            Dim table As New DataTable()
            adapter.Fill(table)

            dgtampil.AllowUserToAddRows = False

            dgtampil.RowTemplate.Height = 80

            ' Mengatur kolom gambar (jika gambar disimpan di database)
            If table.Columns.Contains("Picture") Then ' Ganti dengan nama kolom gambar yang sesuai
                Dim imgc As New DataGridViewImageColumn
                dgtampil.DataSource = table
                imgc = dgtampil.Columns("Picture") ' Ganti dengan nama kolom gambar yang sesuai
                imgc.ImageLayout = DataGridViewImageCellLayout.Stretch ' Menggunakan Zoom agar gambar tidak terdistorsi
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            konn.Close()
        End Try
        FilterData("")
        tidakaktif()
        bersih()

    End Sub

    Private Sub dgtampil_Click(sender As Object, e As EventArgs) Handles dgtampil.Click
        Try
            ' Memastikan bahwa event yang dipicu adalah dari DataGridView
            If TypeOf sender Is DataGridView Then
                Dim dgv As DataGridView = DirectCast(sender, DataGridView)

                ' Mengecek apakah ada baris yang dipilih
                If dgv.CurrentRow IsNot Nothing Then
                    ' Mendapatkan nilai dari baris yang dipilih
                    Dim img As Byte() = DirectCast(dgv.CurrentRow.Cells("Picture").Value, Byte())
                    Dim ms As New MemoryStream(img)
                    pbgambar.Image = Image.FromStream(ms)

                    ' Menampilkan nilai dari sel-sel lainnya
                    txtsku.Text = dgv.CurrentRow.Cells("SKU").Value.ToString()
                    txttitle.Text = dgv.CurrentRow.Cells("Title").Value.ToString()
                    cmbcategory.Text = dgv.CurrentRow.Cells("Category").Value.ToString()
                    txtdeskripsi.Text = dgv.CurrentRow.Cells("Deskripsi").Value.ToString()

                    txtcolour.Text = dgv.CurrentRow.Cells("Colour").Value.ToString()
                    txtmaterials.Text = dgv.CurrentRow.Cells("Materials").Value.ToString()
                    txtprice.Text = dgv.CurrentRow.Cells("Price").Value.ToString()
                    txtstock.Text = dgv.CurrentRow.Cells("Stock").Value.ToString()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub


    Public Sub ExecuteMyQuery(MyQuery As String, MyMessage As String)
        Dim command As New MySqlCommand(MyQuery, konn)

        konn.Open()

        If command.ExecuteNonQuery = 1 Then
            MessageBox.Show(MyMessage)
        Else
            MessageBox.Show("Query Not Executed")
        End If

        konn.Close()
    End Sub

    Private Sub cmdsimpan_Click(sender As Object, e As EventArgs) Handles cmdsimpan.Click
        Try
            Dim checks = {cballsize, cbs, cbm, cbl, cbxl, cbxxl, cbxxxl}
            Dim size As New List(Of String)

            ' Loop melalui semua checkbox dan tambahkan nilai yang dipilih ke dalam list size
            For Each checkbox In checks
                If checkbox.Checked Then
                    size.Add(checkbox.Text)
                End If
            Next

            ' Gabungkan nilai-nilai yang dipilih menjadi satu string dipisahkan koma
            Dim value As String = String.Join(",", size.ToArray())

            ' Simpan gambar ke MemoryStream
            Dim ms As New MemoryStream
            pbgambar.Image.Save(ms, pbgambar.Image.RawFormat)
            Dim img As Byte() = ms.ToArray()

            ' Buka koneksi ke database
            konn.Open()

            ' Gunakan parameterized query untuk menghindari SQL Injection dan masalah dengan data biner
            Dim query As String = "INSERT INTO product (SKU, Title, Category, Deskripsi, Size, Colour, Materials, Price, Stock, Picture) " &
                              "VALUES (@SKu, @Title, @Category, @Deskripsi, @Size, @Colour, @Materials, @Price, @Stock, @Picture)"
            Dim perintah As New MySqlCommand(query, konn)
            perintah.Parameters.AddWithValue("@SKU", txtsku.Text)
            perintah.Parameters.AddWithValue("@Title", txttitle.Text)
            perintah.Parameters.AddWithValue("@Category", cmbcategory.Text)
            perintah.Parameters.AddWithValue("@Deskripsi", txtdeskripsi.Text)
            perintah.Parameters.AddWithValue("@Size", value)
            perintah.Parameters.AddWithValue("@Colour", txtcolour.Text)
            perintah.Parameters.AddWithValue("@Materials", txtmaterials.Text)
            perintah.Parameters.AddWithValue("@Price", txtprice.Text)
            perintah.Parameters.AddWithValue("@Stock", txtstock.Text)
            perintah.Parameters.AddWithValue("@Picture", img)

            ' Eksekusi pernyataan SQL
            perintah.ExecuteNonQuery()

            ' Tutup koneksi ke database
            konn.Close()

            ' Tampilkan pesan berhasil
            MessageBox.Show("Data berhasil disimpan", "Pesan", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cballsize.Checked = False

            cbs.Checked = False
            cbm.Checked = False
            cbl.Checked = False
            cbxl.Checked = False
            cbxxl.Checked = False
            cbxxxl.Checked = False


            ' Bersihkan input dan atur kembali UI
            tampil()
            bersih()

            cmdtambah.Enabled = True
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub cmdupdate_Click(sender As Object, e As EventArgs) Handles cmdupdate.Click
        Try
            Dim checks = {cballsize, cbs, cbm, cbl, cbxl, cbxxl, cbxxxl}
            Dim size As New List(Of String)

            ' Loop melalui semua checkbox dan tambahkan nilai yang dipilih ke dalam list size
            For Each checkbox In checks
                If checkbox.Checked Then
                    size.Add(checkbox.Text)
                End If
            Next

            ' Gabungkan nilai-nilai yang dipilih menjadi satu string dipisahkan koma
            Dim value As String = String.Join(",", size.ToArray())

            Dim ms As New MemoryStream()
            pbgambar.Image.Save(ms, pbgambar.Image.RawFormat)
            Dim img As Byte() = ms.ToArray()

            ' Buka koneksi ke database
            konn.Open()

            ' Gunakan parameterized query
            Dim query As String = "UPDATE product SET Title = @Title, Category = @Category, Deskripsi = @Deskripsi, " &
                              "Size = @Size, Colour = @Colour, Materials = @Materials, Price = @Price, " &
                              "Stock = @Stock, Picture = @Picture WHERE SKU = @SKU"
            Dim perintah As New MySqlCommand(query, konn)
            perintah.Parameters.AddWithValue("@Title", txttitle.Text)
            perintah.Parameters.AddWithValue("@Category", cmbcategory.Text)
            perintah.Parameters.AddWithValue("@Deskripsi", txtdeskripsi.Text)
            perintah.Parameters.AddWithValue("@Size", value)
            perintah.Parameters.AddWithValue("@Colour", txtcolour.Text)
            perintah.Parameters.AddWithValue("@Materials", txtmaterials.Text)
            perintah.Parameters.AddWithValue("@Price", txtprice.Text)
            perintah.Parameters.AddWithValue("@Stock", txtstock.Text)
            perintah.Parameters.AddWithValue("@Picture", img)
            perintah.Parameters.AddWithValue("@SKU", txtsku.Text)

            ' Eksekusi pernyataan SQL
            perintah.ExecuteNonQuery()

            ' Tutup koneksi ke database
            konn.Close()

            ' Tampilkan pesan berhasil
            MessageBox.Show("Data berhasil diupdate", "Pesan", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Refresh tampilan data
            tampil()

            ' Bersihkan input dan atur kembali UI
            bersih()
            cmdtambah.Enabled = True
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub cmdhapus_Click(sender As Object, e As EventArgs) Handles cmdhapus.Click
        konn.Open()
        perintah.Connection = konn
        perintah.CommandType = CommandType.Text
        perintah.CommandText = "delete from product where SKU='" & txtsku.Text & "'"
        perintah.ExecuteNonQuery()
        konn.Close()
        tampil()
        bersih()
    End Sub

    Private Sub txtsku_TextChanged(sender As Object, e As KeyEventArgs) Handles txtsku.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                konn.Open()
                perintah.Connection = konn
                perintah.CommandType = CommandType.Text
                perintah.CommandText = "select * from product " &
                " where sku='" & txtsku.Text & "'"
                cek = perintah.ExecuteReader
                cek.Read()
                MsgBox("data sudah ada...!", MsgBoxStyle.Information, "Pesan")
                If cek.HasRows Then
                    txttitle.Text = cek.Item("Title")
                    cmbcategory.Text = cek.Item("Category")
                    txtdeskripsi.Text = cek.Item("Deskripsi")
                    txtcolour.Text = cek.Item("Colour")
                    txtmaterials.Text = cek.Item("Material")
                    txtprice.Text = cek.Item("Price")
                    txtstock.Text = cek.Item("Stock")
                    pbgambar.Image = cek.Item("Picture")
                    cmdsimpan.Enabled = False
                End If
                konn.Close()
                ' tidakaktif()
                cmdtambah.Enabled = True
        End Select

    End Sub

    Private Sub cmdtambah_Click(sender As Object, e As EventArgs) Handles cmdtambah.Click
        aktif()
        txtsku.Focus()
        cmdtambah.Enabled = False
    End Sub

    Private Sub cmdkeluar_Click(sender As Object, e As EventArgs) Handles cmdkeluar.Click
        Me.Close()
    End Sub

    Private Sub cmdbatal_Click(sender As Object, e As EventArgs) Handles cmdbatal.Click
        End
    End Sub

    Public Sub FilterData(valueToSearch As String)
        Try
            Dim searchQuery As String = "SELECT * FROM product WHERE Title LIKE @search OR Category LIKE @search OR SKU LIKE @search"
            Dim command As New MySqlCommand(searchQuery, konn)
            command.Parameters.AddWithValue("@search", "%" & valueToSearch & "%")

            Dim adapter As New MySqlDataAdapter(command)
            Dim table As New DataTable()
            adapter.Fill(table)

            dgtampil.DataSource = table
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub


    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        FilterData(txtsearch.Text)
    End Sub

    Private Sub pbgambar_Click(sender As Object, e As EventArgs) Handles pbgambar.Click

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
End Class