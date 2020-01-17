Imports System.Data
Imports System.Data.OleDb
Imports System.Speech.Synthesis
Imports System.IO.Ports.SerialPort
Imports HK_Front_End_1.Form1
Imports HK_Front_End_1.Form2
Public Class Form3

    Shared z As Integer

    Shared Sub CallF3(p1 As Integer)
        z = p1
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If z = 1 Then
            Label8.Text = Form1.TextBox1.Text
        Else
            Label8.Text = Form2.TextBox1.Text
        End If
        Dim querry1 As String = "SELECT Age FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry2 As String = "SELECT Gender FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim cmd1 As New OleDbCommand(querry1, con)
        Dim cmd2 As New OleDbCommand(querry2, con)
        con.Open()
        Try
            Label9.Text = cmd1.ExecuteScalar().ToString
            Label10.Text = cmd2.ExecuteScalar().ToString
            con.Close()
        Catch ex As Exception
            MsgBox("Error.", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Shadows Sub Load()
        Label8.Text = Form4.Label22.Text
        Dim querry1 As String = "SELECT Age FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry2 As String = "SELECT Gender FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim cmd1 As New OleDbCommand(querry1, con)
        Dim cmd2 As New OleDbCommand(querry2, con)
        con.Open()
        Try
            Label9.Text = cmd1.ExecuteScalar().ToString
            Label10.Text = cmd2.ExecuteScalar().ToString
            con.Close()
        Catch ex As Exception
            MsgBox("Error.", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Form4.Show()
        If Form4.Visible Then
            Me.Hide()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Panel2.Visible = True
        Panel3.Visible = True
        Panel4.Visible = True
        Panel5.Visible = True
        Panel6.Visible = True
        Panel7.Visible = True
        Panel8.Visible = True
        Panel9.Visible = True
        Panel10.Visible = True
        Button26.Visible = False
        Button27.Visible = True
        RichTextBox4.Visible = False
        Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim querry5 As String = "SELECT [Heart Rate] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry6 As String = "SELECT [Blood Sugar] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry7 As String = "SELECT [Body Temperature] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry8 As String = "SELECT [Blood Pressure] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry9 As String = "SELECT [Prev-Heart Rate] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry10 As String = "SELECT [Prev-Blood Sugar] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry11 As String = "SELECT [Prev-Body Temperature] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry12 As String = "SELECT [Prev-Blood Pressure] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim cmd5 As New OleDbCommand(querry5, con)
        Dim cmd6 As New OleDbCommand(querry6, con)
        Dim cmd7 As New OleDbCommand(querry7, con)
        Dim cmd8 As New OleDbCommand(querry8, con)
        Dim cmd9 As New OleDbCommand(querry9, con)
        Dim cmd10 As New OleDbCommand(querry10, con)
        Dim cmd11 As New OleDbCommand(querry11, con)
        Dim cmd12 As New OleDbCommand(querry12, con)
        con.Open()
        Label36.Text = cmd5.ExecuteScalar().ToString
        Label37.Text = cmd6.ExecuteScalar().ToString
        Label38.Text = cmd7.ExecuteScalar().ToString
        Label39.Text = cmd8.ExecuteScalar().ToString
        Label40.Text = cmd9.ExecuteScalar().ToString
        Label41.Text = cmd10.ExecuteScalar().ToString
        Label42.Text = cmd11.ExecuteScalar().ToString
        Label43.Text = cmd12.ExecuteScalar().ToString
        con.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Panel2.Visible = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        'Panel11.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form1.Show()
        If Form1.Visible Then
            Me.Close()
            Form1.TextBox1.Clear()
            Form1.TextBox2.Clear()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Panel3.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim synth As New SpeechSynthesizer
        synth.Speak("Place your finger in between the L E D and Photodetector for 2 minutes")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel4.Visible = True
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Panel3.Visible = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim synth As New SpeechSynthesizer
        synth.Speak("Place one of your finger in between the N I R sensor and insert another finger in heart beat sensor for about 3 minutes.")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Panel5.Visible = True
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Panel4.Visible = False
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim synth As New SpeechSynthesizer
        synth.Speak("Hold the temperature sensor with 2 fingers for 1 minute")
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Panel6.Visible = True
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Panel5.Visible = False
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim synth As New SpeechSynthesizer
        synth.Speak("Place the cuff on your bare upper arm, one inch above the bend of your elbow. Then pull the end of the cuff so that it's evenly tight around your arm. Finnaly, squeeze the pump for 30 seconds and slowly let out the air.")
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Panel7.Visible = True
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Panel6.Visible = False
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim synth As New SpeechSynthesizer
        synth.Speak("Place the E E G module in your head and move the E E G sensor to your fore-head.")
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Panel8.Visible = True
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Panel7.Visible = False
    End Sub
    
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        MsgBox("Now the appliaction will be connected to the kit.", MessageBoxButtons.OK)
        Panel9.Visible = True
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Panel8.Visible = False
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Panel10.Visible = True
        Button26.Visible = True
        Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim upd1 As String
        upd1 = "UPDATE Table1 SET [Heart Rate]='" + TextBox1.Text + "',[Blood Sugar]='" + TextBox2.Text + "',[Body Temperature]='" + TextBox3.Text + "',[Blood Pressure]='" + TextBox4.Text + "' WHERE [Name]='" + Label8.Text + "'"
        Dim HR, BS, BT, BP As String
        Dim querry1 As String = "SELECT [Heart Rate] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry2 As String = "SELECT [Blood Sugar] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry3 As String = "SELECT [Body Temperature] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry4 As String = "SELECT [Blood Pressure] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim cmd1 As New OleDbCommand(querry1, con)
        Dim cmd2 As New OleDbCommand(querry2, con)
        Dim cmd3 As New OleDbCommand(querry3, con)
        Dim cmd4 As New OleDbCommand(querry4, con)
        Dim cmdupd1 As New OleDbCommand(upd1, con)
        Dim upd2 As String
        If z = 1 Then
            con.Open()
            HR = cmd1.ExecuteScalar().ToString
            BS = cmd2.ExecuteScalar().ToString
            BT = cmd3.ExecuteScalar().ToString
            BP = cmd4.ExecuteScalar().ToString
            cmdupd1.ExecuteNonQuery()
            con.Close()
            upd2 = "UPDATE Table1 SET [Prev-Heart Rate]='" + HR + "',[Prev-Blood Sugar]='" + BS + "',[Prev-Body Temperature]='" + BT + "',[Prev-Blood Pressure]='" + BP + "' WHERE [Name]='" + Label8.Text + "'"
            Dim cmdupd2 As New OleDbCommand(upd2, con)
            con.Open()
            cmdupd2.ExecuteNonQuery()
            con.Close()
        Else
            con.Open()
            HR = cmd1.ExecuteScalar().ToString
            BS = cmd2.ExecuteScalar().ToString
            BT = cmd3.ExecuteScalar().ToString
            BP = cmd4.ExecuteScalar().ToString
            cmdupd1.ExecuteNonQuery()
            con.Close()
            upd2 = "UPDATE Table1 SET [Prev-Heart Rate]='" + HR + "',[Prev-Blood Sugar]='" + BS + "',[Prev-Body Temperature]='" + BT + "',[Prev-Blood Pressure]='" + BP + "' WHERE [Name]='" + Label8.Text + "'"
            Dim cmdupd2 As New OleDbCommand(upd2, con)
            con.Open()
            cmdupd2.ExecuteNonQuery()
            con.Close()
            Button1.Visible = True
            Label2.Visible = True
        End If
        Dim querry5 As String = "SELECT [Heart Rate] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry6 As String = "SELECT [Blood Sugar] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry7 As String = "SELECT [Body Temperature] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry8 As String = "SELECT [Blood Pressure] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry9 As String = "SELECT [Prev-Heart Rate] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry10 As String = "SELECT [Prev-Blood Sugar] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry11 As String = "SELECT [Prev-Body Temperature] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim querry12 As String = "SELECT [Prev-Blood Pressure] FROM Table1 WHERE [Name]= '" & Label8.Text & "';"
        Dim cmd5 As New OleDbCommand(querry5, con)
        Dim cmd6 As New OleDbCommand(querry6, con)
        Dim cmd7 As New OleDbCommand(querry7, con)
        Dim cmd8 As New OleDbCommand(querry8, con)
        Dim cmd9 As New OleDbCommand(querry9, con)
        Dim cmd10 As New OleDbCommand(querry10, con)
        Dim cmd11 As New OleDbCommand(querry11, con)
        Dim cmd12 As New OleDbCommand(querry12, con)
        con.Open()
        Label36.Text = cmd5.ExecuteScalar().ToString
        Label37.Text = cmd6.ExecuteScalar().ToString
        Label38.Text = cmd7.ExecuteScalar().ToString
        Label39.Text = cmd8.ExecuteScalar().ToString
        Label40.Text = cmd9.ExecuteScalar().ToString
        Label41.Text = cmd10.ExecuteScalar().ToString
        Label42.Text = cmd11.ExecuteScalar().ToString
        Label43.Text = cmd12.ExecuteScalar().ToString
        con.Close()
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Panel9.Visible = False
    End Sub

    'Private Sub Button25_Click(sender As Object, e As EventArgs)
    'Panel11.Visible = True
    'End Sub

    Private Sub Button26_Click_1(sender As Object, e As EventArgs) Handles Button26.Click
        Panel10.Visible = False
        Panel9.Visible = False
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Button27.Visible = False
        RichTextBox4.Visible = True
    End Sub

End Class