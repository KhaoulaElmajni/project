Imports System.Data.OleDb

Public Class Form1

    Dim sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\khaoula\Desktop\db\db.mdb"
    Dim sSQL As String
    Dim conn As New OleDbConnection(sConnectionString)
    Dim cmd As New OleDbCommand()
    Dim dr As OleDbDataReader

    Public aaa As String = "estsb"

    Public dt As New DataTable
    Public z As Integer = 0
    Public a As Integer = 0
    Public r As Integer = 0
    Public t As Integer = 0
    Public y As Integer = 0
    Public u As Integer = 0
    Public o As Integer = 0
    Public p As Integer = 0
    Public s As Integer = 0
    Public d As Integer = 0


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        conn.Open()



        If conn.State = ConnectionState.Open Then
            MsgBox("cnx réussie")


        Else
            MsgBox("cnx non réussie")
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Application.Exit()



        'This is the part that you will repear un each query
        'sSQL = "select HM from BULD11 where datedate = '" & DateTimePicker1.Value.ToString().Split(" ")(0) & "'"
        'cmd = New OleDbCommand(sSQL, conn)
        'dr = cmd.ExecuteReader()
        'Do While dr.Read()
        '    MsgBox(dr.Item("HM").ToString)
        'Loop
        'dr.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If CheckBox1.Checked = True Then
            z = 1

        End If
        If CheckBox2.Checked = True Then
            a = 1

        End If
        If CheckBox3.Checked = True Then
            r = 1

        End If
        If CheckBox4.Checked = True Then
            t = 1

        End If
        If CheckBox5.Checked = True Then
            y = 1

        End If
        If CheckBox6.Checked = True Then
            u = 1

        End If
        If CheckBox7.Checked = True Then
            o = 1

        End If
        If CheckBox8.Checked = True Then
            p = 1

        End If
        If CheckBox9.Checked = True Then
            s = 1

        End If
        If CheckBox10.Checked = True Then
            d = 1

        End If

        conn.Close()
        Form2.Show()
        Me.Hide()


    End Sub
End Class
