Imports System.Data.OleDb

Public Class Form2
    Dim sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\khaoula\Desktop\db\db.mdb"
    Dim sSQL As String
    Dim conn As New OleDbConnection(sConnectionString)
    Dim cmd As New OleDbCommand()
    Dim dr As OleDbDataReader


    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        conn.Open()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""

        If Form1.ComboBox1.Text = "D11" Then

            If Form1.z = 1 Then

                'This is the part that you will repear un each query
                sSQL = "select HM from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox1.Text = dr.Item("HM").ToString()

                Loop
                dr.Close()
            End If
            If Form1.a = 1 Then

                sSQL = "select ASUB from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox2.Text = dr.Item("ASUB").ToString()
                Loop
                dr.Close()
            End If
            If Form1.r = 1 Then

                sSQL = "select AEXP from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox3.Text = dr.Item("AEXP").ToString()
                Loop
                dr.Close()
            End If
            If Form1.t = 1 Then

                sSQL = "select AP from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox4.Text = dr.Item("AP").ToString()
                Loop
                dr.Close()
            End If
            If Form1.y = 1 Then

                sSQL = "select TotalJO from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox5.Text = dr.Item("TotalJO").ToString()
                Loop
                dr.Close()
            End If
            If Form1.u = 1 Then

                sSQL = "select TotalAM from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox6.Text = dr.Item("TotalAM").ToString()
                Loop
                dr.Close()
            End If
            If Form1.o = 1 Then

                sSQL = "select TotalFR from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox7.Text = dr.Item("TotalFR").ToString()
                Loop
                dr.Close()
            End If
            If Form1.p = 1 Then

                sSQL = "select Disponibilité from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox8.Text = dr.Item("Disponibilité").ToString()
                Loop
                dr.Close()
            End If
            If Form1.s = 1 Then

                sSQL = "select MTBF from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox9.Text = dr.Item("MTBF").ToString()
                Loop
                dr.Close()
            End If
            If Form1.d = 1 Then

                sSQL = "select MTTR from BULD11 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox10.Text = dr.Item("MTTR").ToString()
                Loop
                dr.Close()
            End If
        End If
        If Form1.ComboBox1.SelectedText = "D9" Then

            If Form1.z = 1 Then

                'This is the part that you will repear un each query
                sSQL = "select HM from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox1.Text = dr.Item("HM").ToString()

                Loop
                dr.Close()
            End If
            If Form1.a = 1 Then

                sSQL = "select ASUB from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox2.Text = dr.Item("ASUB").ToString()
                Loop
                dr.Close()
            End If
            If Form1.r = 1 Then

                sSQL = "select AEXP from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox3.Text = dr.Item("AEXP").ToString()
                Loop
                dr.Close()
            End If
            If Form1.t = 1 Then

                sSQL = "select AP from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox4.Text = dr.Item("AP").ToString()
                Loop
                dr.Close()
            End If
            If Form1.y = 1 Then

                sSQL = "select TotalJO from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox5.Text = dr.Item("TotalJO").ToString()
                Loop
                dr.Close()
            End If
            If Form1.u = 1 Then

                sSQL = "select TotalAM from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox6.Text = dr.Item("TotalAM").ToString()
                Loop
                dr.Close()
            End If
            If Form1.o = 1 Then

                sSQL = "select TotalFR from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox7.Text = dr.Item("TotalFR").ToString()
                Loop
                dr.Close()
            End If
            If Form1.p = 1 Then

                sSQL = "select Disponibilité from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox8.Text = dr.Item("Disponibilité").ToString()
                Loop
                dr.Close()
            End If
            If Form1.s = 1 Then

                sSQL = "select MTBF from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox9.Text = dr.Item("MTBF").ToString()
                Loop
                dr.Close()
            End If
            If Form1.d = 1 Then

                sSQL = "select MTTR from BULLD9 where datedate = '" & Form1.DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
                cmd = New OleDbCommand(sSQL, conn)
                dr = cmd.ExecuteReader()
                Do While dr.Read()
                    TextBox10.Text = dr.Item("MTTR").ToString()
                Loop
                dr.Close()
            End If
        End If
        conn.Close()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.Show()
        Me.Hide()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Application.Exit()
    End Sub
End Class