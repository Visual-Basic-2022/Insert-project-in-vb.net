# Insert-project-in-vb.net
Thanks in advance. I have created a visual basic project using visual studio 2022.
It is running well in dubug mode. But when i create a setup file and install it in th e 
computer and run  the program displaying message "Microsoft Access database Engine requires updatable
query".Please Help me.

Imports System.Data.OleDb
Imports System.Drawing.Text
Imports System.Runtime.CompilerServices.RuntimeHelpers
Public Class Form1
    Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\gutum.mdb")
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Bi_Data()
    End Sub
    Private Sub Bi_Data()
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\gutum.mdb")
        Dim cmd2 As New OleDbCommand("select * from table1 order by id desc", con)
        Dim da As New OleDbDataAdapter
        da.SelectCommand = cmd2
        Dim ta As New DataTable
        ta.Clear()

        Dim v = da.Fill(ta)
        DataGridView1.DataSource = ta
        DataGridView1.Visible = True
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim str As String

        str = "insert into table1(id,wname,wadd) values(@id,@wname,@wadd)"

        Dim cmd1 As New OleDbCommand(str, conn)
        cmd1.Parameters.AddWithValue("@id", TextBox1.Text)
        cmd1.Parameters.AddWithValue("@wname", TextBox2.Text)
        cmd1.Parameters.AddWithValue("@wadd", TextBox3.Text)


        'If conn.State = ConnectionState.Closed Then
        'conn.Open()
        'End If
        'Try

        conn.Open()
            cmd1.ExecuteNonQuery()
        ' Catch ex As OleDbException
        'MessageBox.Show(ex.Message & "_" & ex.Source)
        ' End Try

        conn.Close()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox1.Focus()

        Bi_Data()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub



    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) Then
            e.Handled = True
            MsgBox("only number")
        End If
    End Sub



    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If TextBox1.Text = "" Then
            TextBox1.Focus()
        End If
    End Sub
End Class

