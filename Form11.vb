Public Class Form11

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form1.ComboBox3.Text = TextBox1.Text & "(" & TextBox2.Text & ")"
        Form1.ListView1.Columns(2).Text = Form1.ComboBox3.Text.ToString & " avrg "
        Form1.ListView3.Columns(2).Text = Form1.ComboBox3.Text.ToString
        Form1.TextBox49.Text = "Time VS " + Form1.ComboBox3.Text

        Form1.Chart2.Titles.Add(Form1.TextBox49.Text)
        Form1.TextBox49.Text = "Time VS " + Form1.ComboBox3.Text
        Form1.Chart2.Series(0).Name = Form1.ComboBox3.Text
        Form1.Chart2.ChartAreas(0).AxisX.Title = "Time(HH:MM:SS)"
        Form1.Chart2.ChartAreas(0).AxisY.Title = Form1.ComboBox3.Text
        Form1.ComboBox18.Text = TextBox1.Text
        Me.Close()
    End Sub
End Class