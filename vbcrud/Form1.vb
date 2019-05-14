Imports System.Data.Odbc
Public Class Form1
    Dim Conn As OdbcConnection
    Dim CMD As OdbcCommand
    Dim da As OdbcDataAdapter
    Dim ds As DataSet
    Dim str As String
    Dim id_user As Integer
    Dim nama As String
    Dim alamat As String
    Dim no_hp As String
    Dim jenis_k As String

    Sub TampilData()
        Call Koneksi()
        da = New OdbcDataAdapter("Select * from user", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "user")
        DataGridView1.DataSource = ds.Tables("user")
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "ID User"
            .Columns(1).HeaderCell.Value = "Nama"
            .Columns(2).HeaderCell.Value = "Alamat"
            .Columns(3).HeaderCell.Value = "No. HP"
            .Columns(4).HeaderCell.Value = "Jenis Kelamin"
        End With
    End Sub

    Sub TampilJK()
        ComboBox1.Items.Add("Laki-laki")
        ComboBox1.Items.Add("Perempuan")
    End Sub

    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
    End Sub

    Sub Koneksi()
        str = "Driver={MySQL ODBC 3.51 Driver};database=vbcrud;server=localhost;uid=root;pwd=in12345"
        Conn = New OdbcConnection(str)
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Call TampilData()
        Call TampilJK()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MessageBox.Show("Data belum lengkap!", "Peringatan")
            If TextBox1.Text = "" Then
                TextBox1.Select()
            ElseIf TextBox2.Text = "" Then
                TextBox2.Select()
            ElseIf TextBox3.Text = "" Then
                TextBox3.Select()
            Else
                ComboBox1.Select()
            End If
        Else
            Call Koneksi()
            Dim simpan As String = "insert into user (nama, alamat, no_hp, jenis_k) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
            CMD = New OdbcCommand(simpan, Conn)
            CMD.ExecuteNonQuery()
            MessageBox.Show("Data berhasil ditambahkan", "Info")
            Call TampilData()
            Call KosongkanData()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "Edit" Then
            Button2.Text = "Simpan"
            Button3.Enabled = False
            TextBox1.Text = nama
            TextBox2.Text = alamat
            TextBox3.Text = no_hp
            ComboBox1.Text = jenis_k
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
                MessageBox.Show("Data belum lengkap!", "Peringatan")
                If TextBox1.Text = "" Then
                    TextBox1.Select()
                ElseIf TextBox2.Text = "" Then
                    TextBox2.Select()
                ElseIf TextBox3.Text = "" Then
                    TextBox3.Select()
                Else
                    ComboBox1.Select()
                End If
            Else
                Call Koneksi()
                Dim edit As String = "update user set nama='" & TextBox1.Text & "',alamat='" & TextBox2.Text & "',no_hp='" & TextBox3.Text & "',jenis_k='" & ComboBox1.Text & "' where user_id='" & id_user & "'"
                CMD = New OdbcCommand(edit, Conn)
                CMD.ExecuteNonQuery()
                MessageBox.Show("Data berhasil diupdate", "Info")
                Call TampilData()
                Call KosongkanData()
                Button1.Enabled = True
                Button2.Enabled = False
                Button3.Enabled = False
                Button2.Text = "Edit"
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView1.Rows(e.RowIndex)
            id_user = row.Cells("user_id").Value
            nama = row.Cells("nama").Value.ToString
            alamat = row.Cells("alamat").Value.ToString
            no_hp = row.Cells("no_hp").Value.ToString
            jenis_k = row.Cells("jenis_k").Value.ToString
            If Button2.Text = "Simpan" Then
                TextBox1.Text = nama
                TextBox2.Text = alamat
                TextBox3.Text = no_hp
                ComboBox1.Text = jenis_k
                Button3.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show("Yakin akan dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Call Koneksi()
            Dim hapus As String = "delete From user where user_id='" & id_user & "'"
            CMD = New OdbcCommand(hapus, Conn)
            CMD.ExecuteNonQuery()
            Call TampilData()
            Call KosongkanData()
        End If
    End Sub
End Class