Imports System.Data.OleDb
Public Class pembeli
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String
    Sub tampil()
        dataadapter = New OleDbDataAdapter("select * from Pembeli", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Pembeli")
        DGV.DataSource = (dataset.Tables("Pembeli"))
        DGV.ReadOnly = True
        DGV.Columns(0).Width = 100
        DGV.Columns(1).Width = 100
        DGV.Columns(2).Width = 100
        DGV.Columns(3).Width = 100
    End Sub
    Sub reset()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    Sub kodeotomatis()
        Dim strsementara As String = ""
        Dim strisi As String = ""

        command = New OleDbCommand("select * from Pembeli order by Kodepembeli desc", connection)
        read = command.ExecuteReader()
        If read.Read Then
            strsementara = Mid(read.Item("Kodepembeli"), 2, 2)
            strisi = Val(strsementara) + 1
            TextBox1.Text = "p" + Mid("0", 1, 2 - strisi.Length) & strisi
        Else
            TextBox1.Text = "p01"
        End If
    End Sub
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        home.Show()
        Dispose()
    End Sub

    Private Sub pembeli_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
            command = New OleDbCommand("select * from Pembeli where Kodepembeli='" & TextBox1.Text & "'", connection)
            read = command.ExecuteReader()
            read.Read()
            If Not read.HasRows Then
                Dim dbtambah As String
                dbtambah = "INSERT INTO Pembeli values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                command = New OleDbCommand(dbtambah, connection)
                command.ExecuteNonQuery()
                tampil()
                kodeotomatis()
                reset()
            Else
                Dim dbedit As String
                dbedit = "UPDATE Pembeli SET Namapembeli = '" & TextBox2.Text & "',Alamat = '" & TextBox3.Text & "',Notelp = '" & TextBox4.Text & "' WHERE Kodepembeli = '" & TextBox1.Text & "'"
                command = New OleDbCommand(dbedit, connection)
                command.ExecuteNonQuery()
                tampil()
                kodeotomatis()
                reset()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If MsgBox("Yakin akan menghapus data?", MsgBoxStyle.YesNo, _
"Konfirmasi") = MsgBoxResult.No Then Exit Sub
        Dim dbhapus As String
        dbhapus = "DELETE FROM Pembeli WHERE Kodepembeli = " & _
              "'" & DGV.CurrentRow.Cells(0).Value & "'"

        command = New OleDbCommand(dbhapus, connection)
        command.ExecuteNonQuery()
        tampil()
        kodeotomatis()
        reset()
    End Sub

    Private Sub DGV_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DGV.MouseClick
        TextBox1.Text = DGV.CurrentRow.Cells(0).Value
        TextBox2.Text = DGV.CurrentRow.Cells(1).Value
        TextBox3.Text = DGV.CurrentRow.Cells(2).Value
        TextBox4.Text = DGV.CurrentRow.Cells(3).Value
    End Sub
End Class