Public Class Form12

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        TextBox5.Text = ""
    End Sub

    Private Sub Form12_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label6.Text = Form1.TextBox56.Text
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim i As Integer = ListView1.Items.Count + 1
        ListView1.Items.Add(ListView1.Items.Count + 1)
        ListView1.Items(i - 1).SubItems.Add(TextBox1.Text)
        ListView1.Items(i - 1).SubItems.Add(TextBox2.Text)
        ListView1.Items(i - 1).SubItems.Add(TextBox3.Text)
        ListView1.Items(i - 1).SubItems.Add(ComboBox1.Text)
        If TextBox5.Text = "" Then
            If ComboBox1.Text = "Digital" Then
                ListView1.Items(i - 1).SubItems.Add(2.23)
            ElseIf ComboBox1.Text = "Analog" Then
                ListView1.Items(i - 1).SubItems.Add(1.72)
            End If
        Else
            ListView1.Items(i - 1).SubItems.Add(TextBox5.Text)
        End If

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        TextBox5.Text = ""

        ' Temperature
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim cal As String = 0
        For i = 1 To ListView1.Items.Count
            cal = cal + (CDec(ListView1.Items(i - 1).SubItems(2).Text) * CDec(ListView1.Items(i - 1).SubItems(5).Text)) ^ 2
        Next
        cal = cal + (CDec(Label6.Text)) ^ 2
        TextBox4.Text = Math.Sqrt(cal)
        TextBox6.Text = Val(TextBox11.Text) * CDec(TextBox4.Text)
        TextBox7.Text = Form1.ComboBox3.Text
        TextBox8.Text = Form1.TextBox52.Text
        TextBox9.Text = TextBox6.Text
        TextBox10.Text = ListView1.Items(0).SubItems(3).Text

    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub
End Class