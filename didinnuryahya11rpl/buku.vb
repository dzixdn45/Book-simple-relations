Imports System.Data.OleDb
Public Class buku
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String
    Sub tampil()
        dataadapter = New OleDbDataAdapter("select * from Buku", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Buku")
        DGV.DataSource = (dataset.Tables("Buku"))
        DGV.ReadOnly = True
        DGV.Columns(0).Width = 100
        DGV.Columns(1).Width = 100
        DGV.Columns(2).Width = 100
        DGV.Columns(3).Width = 100
        DGV.Columns(4).Width = 100
        DGV.Columns(5).Width = 100
    End Sub
    Sub reset()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub
    Sub kodeotomatis()
        Dim strsementara As String = ""
        Dim strisi As String = ""

        command = New OleDbCommand("select * from Buku order by kodebuku desc", connection)
        read = command.ExecuteReader()
        If read.Read Then
            strsementara = Mid(read.Item("kodebuku"), 2, 2)
            strisi = Val(strsementara) + 1
            TextBox1.Text = "B" + Mid("0", 1, 2 - strisi.Length) & strisi
        Else
            TextBox1.Text = "B01"
        End If
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didinnuryahya.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        tampil()
        reset()
        kodeotomatis()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Data inputan semua wajib diisi", MsgBoxStyle.OkOnly)
            Exit Sub
        Else
            command = New OleDbCommand("select * from Buku where kodebuku='" & TextBox1.Text & "'", connection)
            read = command.ExecuteReader()
            read.Read()
            If Not read.HasRows Then
                Dim dbtambah As String
                dbtambah = "INSERT INTO Buku values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "')"
                command = New OleDbCommand(dbtambah, connection)
                command.ExecuteNonQuery()
                tampil()
                kodeotomatis()
                reset()
            Else
                Dim dbedit As String
                dbedit = "UPDATE Buku SET judulbuku = '" & TextBox2.Text & "',pengarang = '" & TextBox3.Text & "',penerbit = '" & TextBox4.Text & "',harga = '" & TextBox5.Text & "', jumlah = '" & TextBox6.Text & "' WHERE kodebuku = '" & TextBox1.Text & "'"
                command = New OleDbCommand(dbedit, connection)
                command.ExecuteNonQuery()
                tampil()
                kodeotomatis()
                reset()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If MsgBox("Yakin akan menghapus data?", MsgBoxStyle.YesNo, _
"Konfirmasi") = MsgBoxResult.No Then Exit Sub
        Dim dbhapus As String
        dbhapus = "DELETE FROM Buku WHERE kodebuku = " & _
              "'" & DGV.CurrentRow.Cells(0).Value & "'"

        command = New OleDbCommand(dbhapus, connection)
        command.ExecuteNonQuery()
        tampil()
        kodeotomatis()
        reset()
    End Sub

    Private Sub DGV_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        TextBox1.Text = DGV.CurrentRow.Cells(0).Value
        TextBox2.Text = DGV.CurrentRow.Cells(1).Value
        TextBox3.Text = DGV.CurrentRow.Cells(2).Value
        TextBox4.Text = DGV.CurrentRow.Cells(3).Value
        TextBox5.Text = DGV.CurrentRow.Cells(4).Value
        TextBox6.Text = DGV.CurrentRow.Cells(5).Value
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        home.Show()
        Dispose()
    End Sub
End Class
