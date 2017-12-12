Imports System.Data.OleDb
Public Class transaksi
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String
    Sub buatkolombaru()
        DataGridView1.Columns.Add("kode", "kode")
        DataGridView1.Columns.Add("nama", "nama buku")
        DataGridView1.Columns.Add("harga", "harga")
        DataGridView1.Columns.Add("jumlah", "jumlah")
        DataGridView1.Columns.Add("subtotal", "subtotal")
    End Sub

    Sub aturkolom()
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(4).Width = 100
    End Sub

    Sub totalharga()
        Dim hitungharga As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitungharga = hitungharga + Val(DataGridView1.Rows(i).Cells(4).Value)
            TextBox4.Text = hitungharga
        Next
    End Sub

    Sub tampildaftar()
        command = New OleDbCommand("select * from Pembeli", connection)
        read = command.ExecuteReader
        ComboBox1.Items.Clear()
        Do While read.Read
            ComboBox1.Items.Add(read.Item(0))
        Loop
    End Sub
    Sub nootomatis()
        command = New OleDbCommand("SELECT * FROM transaksi WHERE Kodetransaksi in (select max(Kodetransaksi) from transaksi) order by Kodetransaksi desc", connection)
        read = command.ExecuteReader
        read.Read()
        If Not read.HasRows Then
            TextBox1.Text = Format(Now, "yyMMdd") + "0001"
            TextBox1.ReadOnly = True

        Else
            If Microsoft.VisualBasic.Left(read.GetString(0), 6) <> Format(Now, "yyMMdd") Then
                TextBox1.Text = Format(Now, "yyMMdd") + "0001"
                TextBox1.ReadOnly = True
            Else
                TextBox1.Text = read.GetString(0) + 1
                TextBox1.ReadOnly = True
            End If

        End If
    End Sub
    Sub reset()
        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub
    Sub tampilbuku()
        command = New OleDbCommand("SELECT * FROM Buku", connection)
        read = command.ExecuteReader
        ListBox1.Items.Clear()
        Do While read.Read
            ListBox1.Items.Add(read.Item(0) & Space(5) & read.Item(1))
        Loop
    End Sub
    Private Sub transaksi_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didinnuryahya.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        tampildaftar()
        buatkolombaru()
        aturkolom()
        nootomatis()
        tampilbuku()
        TextBox2.Text = Today
        TextBox4.ReadOnly = True
        TextBox6.ReadOnly = True
        TextBox3.ReadOnly = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        command = New OleDbCommand("select * from Pembeli Where Kodepembeli = '" & ComboBox1.Text & "'", connection)
        read = command.ExecuteReader()
        TextBox3.Clear()
        read.Read()
        If read.HasRows Then
            TextBox3.Text = read.Item("Namapembeli")
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = 0 Then
            command = New OleDbCommand("select * from Buku where kodebuku = '" & DataGridView1.Rows(e.RowIndex).Cells(0).Value & "'", connection)
            read = command.ExecuteReader
            read.Read()
            If read.HasRows Then
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = read.Item(1)
                DataGridView1.Rows(e.RowIndex).Cells(2).Value = read.Item(4)
                DataGridView1.Rows(e.RowIndex).Cells(3).Value = 1
                DataGridView1.Rows(e.RowIndex).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value * DataGridView1.Rows(e.RowIndex).Cells(3).Value
                Call totalharga()

            Else
                MsgBox("Kode obat tidak terdaftar")

            End If
        End If

        If e.ColumnIndex = 3 Then
            command = New OleDbCommand("select * from Buku where kodebuku = '" & DataGridView1.Rows(e.RowIndex).Cells(0).Value & "'", connection)
            Dim rd = command.ExecuteReader
            rd.Read()
            If rd.HasRows Then
                If (DataGridView1.Rows(e.RowIndex).Cells(3).Value > rd.Item(5)) Then
                    MsgBox("Stok Obat hanya ada" & rd.Item(5) & "")
                    DataGridView1.Rows(e.RowIndex).Cells(3).Value = 1
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value * 1
                    Call totalharga()

                Else
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value * DataGridView1.Rows(e.RowIndex).Cells(3).Value
                    Call totalharga()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox5.Text) < Val(TextBox4.Text) Then
                MsgBox("pembayaran anda kurang")
                TextBox6.Text = ""
                TextBox5.Focus()

            Else
                TextBox6.Text = Val(TextBox5.Text) - Val(TextBox4.Text)
                Button1.Focus()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Data anda belum lengkap coeg , tidak ada transaksi atau pembayaran kosong")
            Exit Sub
        End If

        Dim insertresep As String = "insert into transaksi(Kodetransaksi, tgl, kodepembeli, total, dibayar, kembali) values ('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & ComboBox1.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & TextBox6.Text & "')"
        command = New OleDbCommand(insertresep, connection)
        command.ExecuteNonQuery()
        For baris As Integer = 0 To DataGridView1.Rows.Count - 2

            Dim sqlsave As String = "insert into Detailtransaksi(Notransaksi, kodebuku, jumlah, subtotal) values('" & TextBox1.Text & "', '" & DataGridView1.Rows(baris).Cells(0).Value & "', '" & DataGridView1.Rows(baris).Cells(3).Value & "', '" & DataGridView1.Rows(baris).Cells(4).Value & "')"
            command = New OleDbCommand(sqlsave, connection)
            command.ExecuteNonQuery()

            'kurangi stok

            command = New OleDbCommand("select * from Buku where kodebuku='" & DataGridView1.Rows(baris).Cells(0).Value & "'", connection)
            read = command.ExecuteReader
            read.Read()
            If read.HasRows Then
                Dim kuranginstok As String = "update Buku set jumlah= '" & read.Item(5) - DataGridView1.Rows(baris).Cells(3).Value & "' where kodebuku='" & DataGridView1.Rows(baris).Cells(0).Value & "'"
                command = New OleDbCommand(kuranginstok, connection)
                command.ExecuteNonQuery()
            End If
        Next baris
        DataGridView1.Columns.Clear()
        Call reset()
        Call tampilbuku()
        Call buatkolombaru()
        Call nootomatis()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        home.Show()
        Dispose()
    End Sub
End Class