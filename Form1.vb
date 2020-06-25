Imports System.Drawing
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.OleDb
Imports System.Linq
Public Class Form1
    Dim myPort As Array
    Dim baris As Integer
    Dim iterasi As Integer
    Dim TrendLine As New System.Windows.Forms.DataVisualization.Charting.Series("TrendLine")
    Dim vi As Integer
    Dim Rand As New Random
    Dim curve As Integer
    Dim FileName As String
    Dim header As String
    Dim tipeA As Integer = 3
    Dim l As ListViewItem
    Dim P_Stack As String
    Dim N As Integer
    Dim x As String
    Dim ti As String = 0
    Dim ts = 0
    Dim ck As String = 0
    Dim Tc As String
    Dim P_H2O As String
    Dim LV As ListView
    Dim q As Integer
    Dim j_o2 As String
    Dim z As Integer
    Dim v As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim Delimiter As String
    Dim sw As StreamWriter
    Dim legend As String
    Dim sfDialog As New SaveFileDialog
    Dim h As String
    Dim w As String
    Dim sdt As Double = 0
    Dim mn As Double = 0
    Dim k As Integer
    Delegate Sub SetTextCallBack(ByVal [text] As String)

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    
    Public Sub wait(ByVal Dt As Double)
        Dim IDay As Double = Date.Now.DayOfYear
        Dim CDay As Double
        Dim ITime As Double = Date.Now.TimeOfDay.TotalSeconds
        Dim CTime As Double
        Dim DiffDay As Double
        Try
            Do
                Application.DoEvents()
                CDay = Date.Now.DayOfYear
                CTime = Date.Now.TimeOfDay.TotalSeconds
                DiffDay = CDay - IDay
                CTime = CTime + 86400 * DiffDay
                If CTime >= ITime + Dt Then Exit Do
            Loop
        Catch e As Exception
        End Try
    End Sub
  
   
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Try
            With Chart2.ChartAreas(0)
                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            End With
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
     
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ListView1.Enabled = False
            Timer1.Stop()
            Timer2.Stop()
            TextBox63.Enabled = False
            TextBox62.Enabled = False
            Chart2.Series(0).Points.AddXY(0, 0, "0.0000E00")
            Chart2.ChartAreas("ChartArea1").AxisX.Enabled = AxisEnabled.True
            Chart2.ChartAreas("ChartArea1").AxisX.Title = "Time"
            Chart2.ChartAreas("ChartArea1").AxisY.Title = "pH (Acidity)"
            If My.Computer.Clock.LocalTime.Year = 2020 And My.Computer.Clock.LocalTime.Month = 3 And My.Computer.Clock.LocalTime.Day = 30 Then
                Dialog1.Show()
                Me.Visible = False
                Exit Sub
                Me.Visible = True
                Me.Show()
                Me.BringToFront()
            End If
            TextBox61.Visible = False
            
            
                Form5.Hide()
           
            TextBox62.Enabled = False
            TextBox63.Enabled = False

            Timer1.Stop()
            Label104.Text = "Interval (s)"
            TextBox59.Enabled = True
            TextBox60.Enabled = True
           

        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Form1_FormClosing_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            LoginForm1.Close()
            If Application.OpenForms().OfType(Of Form2).Any Then
                Form2.Close()
            ElseIf Application.OpenForms().OfType(Of Form3).Any Then
                Form3.Close()
            ElseIf Application.OpenForms().OfType(Of Form4).Any Then
                Form4.Close()
            ElseIf Application.OpenForms().OfType(Of Form5).Any Then
                Form5.Close()
            ElseIf Application.OpenForms().OfType(Of Form6).Any Then
                Form6.Close()
            ElseIf Application.OpenForms().OfType(Of Form7).Any Then
                Form7.Close()
            ElseIf Application.OpenForms().OfType(Of Form8).Any Then
                Form8.Close()
            ElseIf Application.OpenForms().OfType(Of Form5).Any Then
                Form9.Close()
            ElseIf Application.OpenForms().OfType(Of Form9).Any Then
                Form9.Close()
                Form7.Close()
            Else
                Form10.Close()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub ComboBox33_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            
            c1 = ""
            c2 = ""
            c3 = ""
            c4 = ""
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    
    Public Function ExportListViewToCSV(ByVal filename As String, ByVal lv As ListView) As Boolean
        Try
            Dim os As New StreamWriter(filename)
            For i As Integer = 0 To lv.Columns.Count - 1
                os.Write("""" & lv.Columns(i).Text.Replace("""", """""") & """,")
            Next
            os.WriteLine()
            For i As Integer = 0 To lv.Items.Count - 1
                For j As Integer = 0 To lv.Columns.Count - 1
                    os.Write("""" & lv.Items(i).SubItems(j).Text.Replace("""", """""") + """,")
                Next
                os.WriteLine()
            Next
            os.Close()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
   
    
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Form4.Show()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub BtnScanPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnScanPort.Click
        Try
            CmbScanPort.Items.Clear()
            Dim myPort As Array
            Dim i As Integer
            myPort = IO.Ports.SerialPort.GetPortNames
            CmbScanPort.Items.AddRange(myPort)
            i = CmbScanPort.Items.Count
            i = i - i
            Try
                CmbScanPort.SelectedIndex = i
            Catch ex As Exception
                Dim result As DialogResult
                result = MessageBox.Show("Com Port not detected", "Warning!!!", MessageBoxButtons.OK)
                CmbScanPort.Text = ""
                CmbScanPort.Items.Clear()
                Call Form1_Load(Me, e)
            End Try
            Button28.Enabled = True
            CmbScanPort.DroppedDown = True
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click

        Try
            ' Chart2.Series(0).Points.Clear()
            Timer4.Interval = TextBox17.Text
            Timer3.Start()
            If CmbScanPort.Text = "" Then
                MsgBox("You have to choose Port before measuring!")
                Exit Sub
            End If
            If CmbBaud.Text = "" Then
                MsgBox("You have to choose Baud rate before measuring!")
                Exit Sub
            End If
            If ComboBox2.Text = "Gate Time" Then
                Timer1.Interval = Val(TextBox60.Text) * 1000
                Timer1.Start()
            Else
                Timer1.Interval = Val(TextBox60.Text) * 1000
                Timer1.Start()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Dim DisplaySeriesTrendLine As Boolean = False
    Private Const MAX_RECURSIVE_CALLS As Integer = 1000000000
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        
        If TextBox1.Text = ListView3.Items.Count Then
            wait(1)
            Timer1.Stop()
        Else
            Try
                Dim X_N2 As String = ""
                Dim x_O2 As String
                Dim A As String = ""
                Dim F As String = ""
                Dim R As String = ""
                Dim v As String = ""
                Dim i As Integer = ListView3.Items.Count + 1
                Dim aryText() As String
                Dim aryText1() As String
                With TrendLine
                    .ChartType = SeriesChartType.Line
                    .Color = Color.DodgerBlue
                    .BorderWidth = 1
                    .IsVisibleInLegend = False
                End With
                If SerialPort1.IsOpen Then
                    SerialPort1.Close()
                Else
                    SerialPort1.BaudRate = CmbBaud.SelectedItem
                    SerialPort1.PortName = CmbScanPort.SelectedItem
                    SerialPort1.Open()
                    Dim ka As String = CStr(SerialPort1.ReadLine.ToString)
                    '  wait(1)
                    'If IsNumeric(ka) = True Then
                    'ka = Format(CDbl(ka), "00.00")
                    If Not ka = "" And Not ka = " " Then
                        aryText = ka.Split("|")

                        Label103.Text = CStr(aryText(1))

                        Label111.Text = aryText(2) & "V"
                        Label110.Text = CStr(aryText(3))
                        If IsNumeric(aryText(0)) = True Then
                            If Not Val(aryText(0)) > Val(TextBox62.Text) And Not CStr(aryText(0)) = " " And Not CStr(aryText(0)) = "" And Not Val(aryText(0)) < 0 And Not CStr(ka) = "" And Not Val(aryText(0)) < Val(TextBox63.Text) Then
                                ' If ListView3.Items.Count = Val(TextBox59.Text) Then
                                'TextBox59.Text = i * Val(TextBox59.Text)
                                'ListView1.Items(i - 1).SubItems.Add(CDbl((TextBox61.Text) / Val(TextBox59.Text)))
                                'Else

                                'End If
                                ListView3.Items.Add(ListView3.Items.Count + 1)
                                ListView3.Items(i - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                                aryText1 = aryText(0).Split(".")
                                ListView3.Items(i - 1).SubItems.Add(CDec(aryText1(0) & "." & aryText1(1)))
                                If IsNumeric(aryText(1)) = True Then
                                    ListView3.Items(i - 1).SubItems.Add(CDbl(aryText(1)))
                                Else
                                    ListView3.Items(i - 1).SubItems.Add(0)
                                End If
                                If IsNumeric(aryText(2)) = True Then
                                    ListView3.Items(i - 1).SubItems.Add(CDbl(aryText(2)))
                                Else
                                    ListView3.Items(i - 1).SubItems.Add(0)
                                End If
                                If IsNumeric(aryText(3)) = True Then
                                    ListView3.Items(i - 1).SubItems.Add(CDbl(aryText(3)))
                                Else
                                    ListView3.Items(i - 1).SubItems.Add(0)
                                End If
                                ' 
                                If ListView3.Items(i - 1).SubItems(2).Text > 7.5 Then
                                    Label120.Text = "Basa"
                                ElseIf ListView3.Items(i - 1).SubItems(2).Text > 6.5 And ListView3.Items(i - 1).SubItems(2).Text < 7.5 Then
                                    Label120.Text = "Normal"
                                Else
                                    Label120.Text = "Asam"
                                End If
                                TextBox61.Text = CDbl(TextBox61.Text) + CDbl(ListView3.Items(i - 1).SubItems(2).Text)
                                TextBox3.Text = CDbl(TextBox3.Text) + CDbl(ListView3.Items(i - 1).SubItems(3).Text)
                                TextBox5.Text = CDbl(TextBox5.Text) + CDbl(ListView3.Items(i - 1).SubItems(4).Text)
                                TextBox6.Text = CDbl(TextBox6.Text) + CDbl(ListView3.Items(i - 1).SubItems(5).Text)
                                Dim avrg As String
                                Dim avrgasint As Integer
                                If ListView1.Enabled = True Then
                                    avrg = ListView3.Items.Count / TextBox59.Text
                                    If Integer.TryParse(avrg, avrgasint) Then
                                        ListView1.Items.Add(ListView1.Items.Count + 1)
                                        ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                                        'aryText1 = aryText(0).Split(".")
                                        ListView1.Items(ListView1.Items.Count - 1).SubItems.Add((CDbl(TextBox61.Text / CDbl(TextBox59.Text))))
                                        ListView1.Items(ListView1.Items.Count - 1).SubItems(2).Text = Format(CDbl(ListView1.Items(ListView1.Items.Count - 1).SubItems(2).Text), "0.00")
                                        ' MsgBox(TextBox61.Text & " " & (10 * (ListView1.Items.Count)))
                                        ListView1.Items(ListView1.Items.Count - 1).SubItems.Add((CDbl(TextBox3.Text / CDbl(TextBox59.Text))))
                                        ListView1.Items(ListView1.Items.Count - 1).SubItems.Add((CDbl(TextBox5.Text / CDbl(TextBox59.Text))))
                                        ListView1.Items(ListView1.Items.Count - 1).SubItems.Add((CDbl(TextBox6.Text / CDbl(TextBox59.Text))))
                                        TextBox61.Text = "0"
                                        TextBox3.Text = "0"
                                        TextBox5.Text = "0"
                                        TextBox6.Text = "0"
                                        If ListView1.Items.Count >= 2 Then
                                            ListView1.Items(ListView1.Items.Count - 1).SubItems(6).Text = CDec(ListView1.Items(ListView1.Items.Count - 1).SubItems(2).Text) - CDec(ListView1.Items(ListView1.Items.Count - 2).SubItems(2).Text)
                                        Else
                                            ListView1.Items(ListView1.Items.Count - 1).SubItems(6).Text = "0"
                                        End If
                                    End If

                                End If

                                'ListView3.Items.Clear()
                                'TextBox61.Text = 0
                                If i >= 2 Then
                                    ListView3.Items(i - 1).SubItems.Add(CDec(ListView3.Items(i - 1).SubItems(2).Text) - CDec(ListView3.Items(i - 2).SubItems(2).Text))
                                Else
                                    ListView3.Items(i - 1).SubItems.Add(0)
                                End If
                                If ComboBox4.Text = "Point" Then
                                    Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                                ElseIf ComboBox4.Text = "Area" Then
                                    Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                                ElseIf ComboBox4.Text = "Fast Line" Then
                                    Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                                Else
                                    Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                                End If
                                'Chart2.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
                                'Chart2.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                                'Chart2.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                                If Not aryText(0) = "" And Not aryText(0) = " " Then
                                    Chart2.Series(0).Points.AddXY((ListView3.Items(i - 1).SubItems(1).Text.ToString), CDbl(aryText1(0) & "." & aryText1(1)))
                                Else
                                    wait(1)
                                End If

                                If ListView3.Items(i - 1).BackColor = Color.Red Then
                                    Chart2.Series(0).Points(i - 1).Label = "Contain Error Data"
                                    Chart2.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                                    Chart2.Series(0).Points(i - 1).MarkerSize = 10
                                    Chart2.Series(0).Points(i - 1).MarkerColor = Color.Red
                                End If
                                ListView3.Items(i - 1).EnsureVisible()
                                Chart2.ChartAreas("ChartArea1").AxisX.ScrollBar.Size = 10
                                Chart2.ChartAreas("ChartArea1").AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll
                                Chart2.ChartAreas("ChartArea1").AxisX.ScrollBar.IsPositionedInside = True
                                Chart2.ChartAreas("ChartArea1").AxisX.ScrollBar.Enabled = True
                                If ComboBox5.Text = "Red" Then
                                    Chart2.Series(0).Color = Color.Red
                                ElseIf ComboBox5.Text = "Green" Then
                                    Chart2.Series(0).Color = Color.Green
                                ElseIf ComboBox5.Text = "Blue" Then
                                    Chart2.Series(0).Color = Color.Blue
                                Else
                                    Chart2.Series(0).Color = Color.Brown
                                End If
                                If ComboBox6.Text = "Dash" Then
                                    With Chart2.ChartAreas(0)
                                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                    End With
                                Else
                                    With Chart2.ChartAreas(0)
                                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                    End With
                                End If
                                Chart2.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                                Chart2.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Left
                                If TextBox55.Text = "" Then
                                    TextBox55.Text = CDec(ListView3.Items(0).SubItems(2).Text)
                                Else
                                    If CDec(ListView3.Items(i - 1).SubItems(2).Text.ToString) > CDec(TextBox55.Text) And Not ListView3.Items(i - 1).SubItems(2).Text.ToString = "" And Not TextBox55.Text = "" Then
                                        TextBox55.Text = CDec(ListView3.Items(i - 1).SubItems(2).Text.ToString)
                                        ListView3.Items(i - 1).BackColor = Color.Yellow
                                        'Chart2.ChartAreas("ChartArea1").AxisY.Maximum = CDec(TextBox55.Text) + CDec(TextBox57.Text / 8)
                                        '  For i = 1 To ListView3.Items.Count - 1
                                        'If Chart2.Series(0).Points(i - 1).MarkerColor = Color.Yellow Then
                                        'Chart2.Series(0).Points(i - 1).MarkerSize = 0
                                        'Chart2.Series(0).Points(i - 1).Label = ""
                                        ' End If
                                        '  Next
                                        '   Chart2.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                                        '   Chart2.Series(0).Points(i - 1).MarkerSize = 10
                                        '  Chart2.Series(0).Points(i - 1).MarkerColor = Color.Yellow
                                        '  Chart2.Series(0).Points(i - 1).LabelForeColor = Color.Blue
                                        ' Chart2.Series(0).Points(i - 1).Label = ListView3.Items(i - 1).SubItems(2).Text
                                        Label100.Text = Val(Label100.Text) + 1
                                    Else
                                        'TextBox55.Text = Format(CDbl(TextBox55.Text), "0.0000E00")
                                        'Chart2.ChartAreas("ChartArea1").AxisY.Maximum = CDec(TextBox55.Text) + CDec(TextBox57.Text / 8)
                                    End If
                                End If
                                If TextBox54.Text = "" Then
                                    TextBox54.Text = (CDec(ListView3.Items(0).SubItems(2).Text))
                                Else
                                    If CDec(ListView3.Items(i - 1).SubItems(2).Text.ToString) < CDec(TextBox54.Text) Then
                                        TextBox54.Text = CDbl(ListView3.Items(i - 1).SubItems(2).Text.ToString)
                                        ' Chart2.ChartAreas("ChartArea1").AxisY.Minimum = CDec(TextBox54.Text) - CDec(TextBox57.Text / 8)

                                        ' Chart2.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                                        ' Chart2.Series(0).Points(i - 1).MarkerSize = 10
                                        ' Chart2.Series(0).Points(i - 1).MarkerColor = Color.LightGreen
                                        ' Chart2.Series(0).Points(i - 1).LabelForeColor = Color.Blue
                                        ' Chart2.Series(0).Points(i - 1).Label = ListView3.Items(i - 1).SubItems(2).Text
                                        ' Label101.Text = Val(Label101.Text) + 1
                                    Else
                                        TextBox54.Text = CDec(TextBox54.Text)
                                        'Chart2.ChartAreas("ChartArea1").AxisY.Minimum = CDec(TextBox54.Text) - (TextBox57.Text / 8)
                                    End If
                                End If
                                If Not TextBox52.Text = "" Then
                                    TextBox52.Text = Format(CDbl((TextBox52.Text * (ListView3.Items.Count - 1) + (CDec(ListView3.Items(i - 1).SubItems(2).Text))) / ListView3.Items.Count), "0.00E00")
                                Else
                                    TextBox52.Text = Format(CDbl(ListView3.Items(i - 1).SubItems(2).Text), "0.000E00")
                                End If
                                If Not TextBox56.Text = "" Then
                                    TextBox56.Text = Format(CDbl(Math.Sqrt(((TextBox56.Text ^ 2) * (ListView3.Items.Count - 2) / (ListView3.Items.Count - 1)) + ((Math.Abs(CDec(ListView3.Items(i - 1).SubItems(2).Text) - CDec(TextBox52.Text))) ^ 2 / (ListView3.Items.Count - 1)))), "0.00E00")
                                Else
                                    TextBox56.Text = 0
                                End If
                                If i > 1 And Not TextBox57.Text = "" Then
                                    TextBox57.Text = CDec(TextBox55.Text - TextBox54.Text)
                                Else
                                    TextBox57.Text = 0
                                End If
                                wait(TextBox60.Text)
                                ' Else
                                ' vi = ListView3.Items.Count + 1
                                ' ListView3.Items.Add(ListView3.Items.Count + 1)
                                ' ListView3.Items(vi - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                                ' ListView3.Items(vi - 1).SubItems.Add(aryText(0))
                                ' TextBox61.Text = CDec(TextBox61.Text) + CDec(ListView3.Items(vi - 1).SubItems(2).Text)
                                'ProgressBar1.Increment((100 / TextBox59.Text) + 1)
                                ' If ProgressBar1.Value = 100 Then
                                '  Label106.ForeColor = Color.Red
                                '  Label106.Text = "Data ready"
                                '  ProgressBar1.ForeColor = Color.Red
                                '  ProgressBar1.Value = 0
                                ' Else
                                ' Label106.Text = "Read Data"
                                ' Label106.ForeColor = Color.Blue
                                '  End If
                                ' End If
                            Else
                                Label106.ForeColor = Color.Red
                                Label106.BackColor = Color.Transparent
                                Label106.Text = "There is Data Error, Wait.."
                                'ProgressBar1.ForeColor = Color.Red
                            End If
                        Else
                            wait(1)
                        End If
                    End If
                    SerialPort1.Close()
                    'wait(0.01)
                    TextBox53.Text = TextBox53.Text + 1

                    End If


                    'TextBox51.Text = TextBox51.Text + Val(TextBox60.Text)
                    TextBox50.Text = ((ListView3.Items.Count)) - TextBox64.Text
                    If TextBox73.Text = TextBox50.Text + TextBox53.Text Then
                        Button26.PerformClick()
                        MsgBox("You reach Maximum Measurements")
                        Exit Sub
                    End If
                    If ListView3.Items.Count > 1 Then
                        'TextBox55.Text = Format(CDec(TextBox55.Text), "0.00E00")
                        'TextBox54.Text = Format(CDec(TextBox54.Text), "0.00E00")
                        'TextBox52.Text = Format(CDec(TextBox52.Text), "0.00E00")
                        'TextBox56.Text = Format(CDec(TextBox56.Text), "0.00E00")
                        'TextBox57.Text = Format(CDec(TextBox57.Text), "0.00E00")
                    End If
            Catch s As Exception
                ' MsgBox("Data Error! continiue to next run?")
                'Timer1.Stop()
                '  Exit Sub
            End Try

            If TextBox55.Text = "" Then
                TextBox55.Text = 0
            ElseIf TextBox54.Text = "" Then
                TextBox54.Text = 0
            Else
                Chart2.ChartAreas("ChartArea1").AxisY.Maximum = CDec(TextBox55.Text) + TextBox2.Text 'CDec(TextBox55.Text - TextBox54.Text) / 5
                Chart2.ChartAreas("ChartArea1").AxisY.Minimum = CDec(TextBox54.Text) - TextBox2.Text 'Double.NaN
                Chart2.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                Chart2.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
            End If
           
        End If
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            Button26.PerformClick()
            'TextBox47.Text = ""
            legend = ""
            SerialPort1.Close()
            TextBox50.Text = 0
            Timer3.Start()
            ' TextBox51.Text = 0
            TextBox52.Text = 0
            TextBox53.Text = 0
            TextBox54.Text = ""
            TextBox55.Text = ""
            TextBox56.Text = 0
            TextBox57.Text = 0
            'ProgressBar1.Value = 0
            Label106.Text = "Standby"
            Me.Chart2.Series(0).Points.Clear()
            ListView3.Items.Clear()
            ListView3.Items.Clear()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Try
            GroupBox17.Visible = True
            GroupBox17.Enabled = True
            SerialPort1.Close()
            Timer1.Stop()
            Timer3.Stop()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Try
            Dim SaveFile As New SaveFileDialog
            SaveFile.FileName = ""
            SaveFile.Filter = "Text Files (*.txt)|*.txt"
            SaveFile.Title = "Save"
            SaveFile.ShowDialog()
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)
            Dim col As ColumnHeader
            Dim columnnames As String = ""
            For Each col In ListView3.Columns
                If String.IsNullOrEmpty(columnnames) Then
                    columnnames = col.Text
                Else
                    columnnames &= "|" & col.Text
                End If
            Next
            Write.Write(columnnames & vbCrLf)
            For Me.baris = 1 To ListView3.Items.Count - 1
                Write.Write(ListView3.Items(baris - 1).SubItems(0).Text & "|" & ListView3.Items(baris - 1).SubItems(1).Text & "|" & ListView3.Items(baris - 1).SubItems(2).Text & "|" & ListView3.Items(baris - 1).SubItems(3).Text & "|" & ListView3.Items(baris - 1).SubItems(4).Text & "|" & ListView3.Items(baris - 1).SubItems(5).Text & "|" & ListView3.Items(baris - 1).SubItems(6).Text & vbCrLf)
            Next baris
            Write.Close()
        Catch d As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Try
            Chart2.ChartAreas(0).AxisY.ScaleView.Size = Math.Abs(CDec(TextBox55.Text) - CDec(TextBox54.Text)) * 2
            Chart2.ChartAreas(0).AxisX.ScaleView.Size = Math.Abs(CDec(ListView3.Items(ListView3.Items.Count).SubItems(2).Text.ToString) - CDec(ListView3.Items(0).SubItems(2).Text.ToString)) * 2
            Chart2.ChartAreas(0).AxisX.ScrollBar.Enabled = True
            Chart2.ChartAreas(0).AxisY.ScrollBar.Enabled = True
            Chart2.ChartAreas(0).CursorX.IsUserEnabled = True
            Chart2.ChartAreas(0).CursorY.IsUserEnabled = True
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            TextBox62.Enabled = True
            TextBox63.Enabled = True
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            If ComboBox2.Text = "Interval Time" Then
                Label104.Text = "Interval (s)"
                TextBox59.Enabled = True
                TextBox60.Enabled = True
            ElseIf ComboBox2.Text = "Gate Time" Then
                TextBox60.Enabled = True
                TextBox59.Enabled = False
                Label104.Text = "Gate Time (s)"
            Else
                TextBox60.Enabled = False
                TextBox59.Enabled = True
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Chart2.Titles.Clear()
        Try
            If ComboBox3.Text = "Custom" Then
                Form11.Show()
            Else
                ListView1.Columns(2).Text = ComboBox3.Text.ToString & " avrg "
                ListView3.Columns(2).Text = ComboBox3.Text.ToString

                TextBox49.Text = "Time VS " + ComboBox3.Text
                Chart2.Titles.Add(TextBox49.Text)
                'Chart2.Series(0).Name = ComboBox3.Text
                Chart2.ChartAreas(0).AxisX.Title = "Time(HH:MM:SS)"
                Chart2.ChartAreas(0).AxisY.Title = ComboBox3.Text
            End If

        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Try
            Chart2.ChartAreas(0).AxisY.ScaleView.Size = Math.Abs(CDec(TextBox55.Text) - CDec(TextBox54.Text)) / 2
            Chart2.ChartAreas(0).AxisX.ScaleView.Size = Math.Abs(CDec(ListView3.Items(ListView3.Items.Count).SubItems(2).Text.ToString) - CDec(ListView3.Items(0).SubItems(2).Text.ToString)) / 2
            Chart2.ChartAreas(0).AxisX.ScrollBar.Enabled = True
            Chart2.ChartAreas(0).AxisY.ScrollBar.Enabled = True
            Chart2.ChartAreas(0).CursorX.IsUserEnabled = True
            Chart2.ChartAreas(0).CursorY.IsUserEnabled = True
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Try
            With Chart2.ChartAreas(0)
                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            End With
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Try
            With Chart2.ChartAreas(0)
                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            End With
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim i As Integer
        Try
            Chart2.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Left
            If TextBox55.Text = "" Then
                TextBox55.Text = ListView3.Items(0).SubItems(2).Text
            Else
                If CDec(ListView3.Items(i - 1).SubItems(2).Text.ToString) > CDec(TextBox55.Text) And Not ListView3.Items(i - 1).SubItems(2).Text.ToString = "" And Not TextBox55.Text = "" Then
                    TextBox55.Text = CDec(ListView3.Items(i - 1).SubItems(2).Text.ToString)
                    ListView3.Items(i - 1).BackColor = Color.Yellow
                    Chart2.ChartAreas("ChartArea1").AxisY.Maximum = CDec(TextBox55.Text) + (TextBox57.Text / 8)
                    Chart2.Series(0).Points(i - 1).Label = "max(" & Label100.Text + 1 & ")= " + CDec(ListView3.Items(i - 1).SubItems(2).Text)
                    Chart2.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                    Chart2.Series(0).Points(i - 1).MarkerSize = 10
                    Chart2.Series(0).Points(i - 1).MarkerColor = Color.Yellow
                    Label100.Text = Val(Label100.Text) + 1
                Else
                    TextBox55.Text = CDec(TextBox55.Text)
                    Chart2.ChartAreas("ChartArea1").AxisY.Maximum = CDec(TextBox55.Text) + CDec(TextBox57.Text / 8)
                End If
            End If
            If TextBox54.Text = "" Then
                TextBox54.Text = ListView3.Items(0).SubItems(2).Text
            Else
                If CDec(ListView3.Items(i - 1).SubItems(2).Text.ToString) < CDec(TextBox54.Text) Then
                    TextBox54.Text = Format(CDec(ListView3.Items(i - 1).SubItems(2).Text.ToString), "00.00")
                    Chart2.ChartAreas("ChartArea1").AxisY.Minimum = CDec(TextBox54.Text) - (TextBox57.Text / 8)
                    ListView3.Items(i - 1).BackColor = Color.LightGreen
                    Chart2.Series(0).Points(i - 1).Label = "min(" & Label101.Text + 1 & ")= " + (ListView3.Items(i - 1).SubItems(2).Text)
                    Chart2.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                    Chart2.Series(0).Points(i - 1).MarkerSize = 10
                    Chart2.Series(0).Points(i - 1).MarkerColor = Color.LightGreen
                    Label101.Text = Val(Label101.Text) + 1
                Else
                    TextBox54.Text = CDec(TextBox54.Text)
                    Chart2.ChartAreas("ChartArea1").AxisY.Minimum = CDec(TextBox54.Text) - (TextBox57.Text / 8)
                End If
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Try
            With Chart2.ChartAreas(0)
                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            End With
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Try
            With Chart2.ChartAreas(0)
                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            End With
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            For g = 1 To ListView3.Items.Count
                Chart2.Series(CInt(curve)).Points(g - 1).Label = ""
                Chart2.Series(CInt(curve)).Points(g - 1).MarkerSize = 0
            Next g
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Try
            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|SVG (*.svg)|*.svg|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
            saveFileDialog1.FilterIndex = 2
            saveFileDialog1.RestoreDirectory = True
            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                Dim format As ChartImageFormat = ChartImageFormat.Bmp
                If saveFileDialog1.FileName.EndsWith("bmp") Then
                    format = ChartImageFormat.Bmp
                Else
                    If saveFileDialog1.FileName.EndsWith("jpg") Then
                        format = ChartImageFormat.Jpeg
                    Else
                        If saveFileDialog1.FileName.EndsWith("emf") Then
                            format = ChartImageFormat.Emf
                        Else
                            If saveFileDialog1.FileName.EndsWith("gif") Then
                                format = ChartImageFormat.Gif
                            Else
                                If saveFileDialog1.FileName.EndsWith("png") Then
                                    format = ChartImageFormat.Png
                                Else
                                    If saveFileDialog1.FileName.EndsWith("tif") Then
                                        format = ChartImageFormat.Tiff
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                Chart2.SaveImage(saveFileDialog1.FileName, format)
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            Chart2.Series(0).Points.Clear()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            TabControl1.SelectedIndex = 2
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        Try
            TabControl1.SelectedIndex = 1
            BtnScanPort.PerformClick()

        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
   
    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Form6.Show()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    
    Private Sub TextBox73_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox73.TextChanged
        If TextBox73.Text > 1000000000 Then
            MsgBox("Maximum data must be bellow 1000000000!")
            TextBox73.Text = 1000000000
            Exit Sub
        End If
    End Sub

    Private Sub ProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox55_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox55.TextChanged
        For g = 1 To ListView3.Items.Count - 1
            If Chart2.Series(0).Points(g - 1).MarkerColor = Color.Yellow Then
                Chart2.Series(0).Points(g - 1).MarkerSize = 0
                Chart2.Series(0).Points(g - 1).Label = ""
            End If
        Next
        'Dim i As Integer = ListView3.Items.Count
        'Chart2.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
        'Chart2.Series(0).Points(i - 1).MarkerSize = 10
        'Chart2.Series(0).Points(i - 1).MarkerColor = Color.Yellow
        'Chart2.Series(0).Points(i - 1).LabelForeColor = Color.Blue
        'Chart2.Series(0).Points(i - 1).Label = TextBox55.Text
        'TextBox55.Text = Format(CDec(TextBox55.Text), "0.00000")
        For g = 1 To ListView3.Items.Count
            If ListView3.Items(g - 1).SubItems(2).Text = TextBox55.Text Then
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                Chart2.Series(0).Points(g - 1).LabelForeColor = Color.Blue
                Chart2.Series(0).Points(g - 1).Label = TextBox55.Text
            End If
        Next

    End Sub

    Private Sub TextBox54_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox54.TextChanged
        For i = 1 To ListView3.Items.Count - 1
            If Chart2.Series(0).Points(i - 1).MarkerColor = Color.LightGreen Then
                Chart2.Series(0).Points(i - 1).MarkerSize = 0
                Chart2.Series(0).Points(i - 1).Label = ""
            End If
        Next
        '   Dim j As Integer = ListView3.Items.Count
        '  Chart2.Series(0).Points(j - 1).MarkerStyle = MarkerStyle.Circle
        '  Chart2.Series(0).Points(j - 1).MarkerSize = 10
        '  Chart2.Series(0).Points(j - 1).MarkerColor = Color.LightGreen
        '  Chart2.Series(0).Points(j - 1).LabelForeColor = Color.Blue
        ' Chart2.Series(0).Points(j - 1).Label = TextBox54.Text
        For g = 1 To ListView3.Items.Count
            If ListView3.Items(g - 1).SubItems(2).Text = TextBox54.Text Then
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                Chart2.Series(0).Points(g - 1).LabelForeColor = Color.Blue
                Chart2.Series(0).Points(g - 1).Label = TextBox54.Text
            End If
        Next
        'TextBox54.Text = Format(CDec(TextBox54.Text), "0.00000")
    End Sub

    Private Sub Chart2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub CmbScanPort_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbScanPort.SelectedIndexChanged

    End Sub

    Private Sub Label103_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label103.Click

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        'If ComboBox2.Text = "Interval Time" Then
        If SerialPort1.IsOpen Then
            SerialPort1.Close()
        Else
            SerialPort1.BaudRate = CmbBaud.SelectedItem
            SerialPort1.PortName = CmbScanPort.SelectedItem
            SerialPort1.Open()
            Dim ka1 As String = CStr(SerialPort1.ReadLine.ToString)
            Dim aryText() As String
            'If IsNumeric(ka) = True Then
            'ka = Format(CDbl(ka), "00.00")
            aryText = ka1.Split("|")
            'If Not Val(aryText(0) ) > Val(TextBox62.Text) And Not CStr(aryText(0) ) = "" And Not Val(aryText(0)) < 2 And Not CStr(k) = "" And Not Val(aryText(0)) < Val(TextBox63.Text) Then
            Label111.Text = aryText(1)
            Label103.Text = aryText(2)
            Label110.Text = aryText(3)
            'End If
        End If

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Try
            Dim SaveFile As New SaveFileDialog
            SaveFile.FileName = ""
            SaveFile.Filter = "Text Files (*.txt)|*.txt"
            SaveFile.Title = "Save"
            SaveFile.ShowDialog()
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)
            Dim col As ColumnHeader
            Dim columnnames As String = ""
            ' For Each col In ListView3.Columns
            'If String.IsNullOrEmpty(columnnames) Then
            'columnnames = col.Text
            'Else
            ' columnnames &= "|" & col.Text
            ' End If
            ' Next
            ' Write.Write(columnnames & vbCrLf)
            ' For Me.baris = 1 To ListView3.Items.Count - 1
            Write.Write(Label97.Text & "=" & TextBox55.Text & vbCrLf)
            Write.Write(Label96.Text & "=" & TextBox54.Text & vbCrLf)
            Write.Write(Label91.Text & "=" & TextBox52.Text & vbCrLf)
            Write.Write(Label94.Text & "=" & TextBox56.Text & vbCrLf)
            Write.Write(Label95.Text & "=" & TextBox57.Text & vbCrLf)
            Write.Write(Label92.Text & "=" & TextBox51.Text & vbCrLf)
            Write.Write(Label93.Text & "=" & TextBox50.Text & vbCrLf)
            Write.Write(Label90.Text & "=" & TextBox53.Text & vbCrLf)
            ' Write.Write(Label118.Text & "=" & Label117.Text & vbCrLf)
            ' Next baris
            Write.Close()
        Catch d As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub TabPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        ListView1.Enabled = True
    End Sub

    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        Try
            Dim SaveFile As New SaveFileDialog
            SaveFile.FileName = ""
            SaveFile.Filter = "Text Files (*.txt)|*.txt"
            SaveFile.Title = "Save"
            SaveFile.ShowDialog()
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)
            Dim col As ColumnHeader
            Dim columnnames As String = ""
            For Each col In ListView1.Columns
                If String.IsNullOrEmpty(columnnames) Then
                    columnnames = col.Text
                Else
                    columnnames &= "|" & col.Text
                End If
            Next
            Write.Write(columnnames & vbCrLf)
            For Me.baris = 1 To ListView1.Items.Count - 1
                Write.Write(ListView1.Items(baris - 1).SubItems(0).Text & "|" & ListView1.Items(baris - 1).SubItems(1).Text & "|" & ListView1.Items(baris - 1).SubItems(2).Text & "|" & ListView1.Items(baris - 1).SubItems(3).Text & "|" & ListView1.Items(baris - 1).SubItems(4).Text & "|" & ListView1.Items(baris - 1).SubItems(5).Text & "|" & ListView1.Items(baris - 1).SubItems(6).Text & vbCrLf)
            Next baris
            Write.Close()
        Catch d As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        ListView1.Enabled = False
    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        ListView1.Items.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox55.Text = ""
        TextBox54.Text = ""
        TextBox52.Text = ""
        TextBox56.Text = ""
        TextBox57.Text = ""
        TextBox51.Text = ""
        TextBox50.Text = ""
        TextBox53.Text = ""
    End Sub

    Private Sub Button55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button55.Click
        Form12.Show()
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button56.Click
        Button69.PerformClick()
        '   If Not SeriesChartType.Point = True Then
        wait(1)
        'Dim i As Integer = ListView3.Items.Count
        'Chart2.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
        'Chart2.Series(0).Points(i - 1).MarkerSize = 10
        'Chart2.Series(0).Points(i - 1).MarkerColor = Color.Yellow
        'Chart2.Series(0).Points(i - 1).LabelForeColor = Color.Blue
        'Chart2.Series(0).Points(i - 1).Label = TextBox55.Text
        'TextBox55.Text = Format(CDec(TextBox55.Text), "0.00000")
        For g = 1 To Chart2.Series(0).Points.Count
            If CDbl(ListView3.Items(g - 1).SubItems(2).Text) = CDbl(TextBox55.Text) Then
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                Chart2.Series(0).Points(g - 1).LabelForeColor = Color.Blue
                Chart2.Series(0).Points(g - 1).Label = TextBox55.Text
            End If
        Next
        wait(1)
        For g = 1 To Chart2.Series(0).Points.Count
            If CDbl(ListView3.Items(g - 1).SubItems(2).Text) = CDbl(TextBox54.Text) Then
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                Chart2.Series(0).Points(g - 1).LabelForeColor = Color.Blue
                Chart2.Series(0).Points(g - 1).Label = TextBox54.Text
            End If
        Next

    End Sub

    Private Sub Button69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button69.Click
        Chart2.Series(0).Points.Clear()
        For i = 1 To ListView3.Items.Count
            ListView3.Items.Add(ListView3.Items.Count + 1)
            ListView3.Items(i - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
            ListView3.Items(i - 1).SubItems.Add(CDec(ListView3.Items(i - 1).SubItems(2).Text))
        Next
        Chart2.ChartAreas("ChartArea1").AxisY.Maximum = CDec(TextBox55.Text) + TextBox2.Text 'CDec(TextBox55.Text - TextBox54.Text) / 5
        Chart2.ChartAreas("ChartArea1").AxisY.Minimum = CDec(TextBox54.Text) - TextBox2.Text 'Double.NaN
        Chart2.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart2.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Label117.Text = CDbl(Label117.Text) + CDbl(TextBox60.Text)
        TextBox51.Text = Label117.Text
    End Sub
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        TextBox8.Enabled = False
        TextBox10.Enabled = True
    End Sub

    Private Sub Button70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button70.Click
        Timer4.Enabled = True
        Timer4.Interval = TextBox17.Text * 1000
        Timer4.Start()
    End Sub

    Private Sub Button71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button71.Click
        Timer4.Stop()
        Timer6.Stop()
    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        Try
            TextBox21.Text = Val(TextBox21.Text) + 1
            Dim i As Integer = ListView8.Items.Count + 1

            With TrendLine
                .ChartType = SeriesChartType.Line
                .Color = Color.DodgerBlue
                .BorderWidth = 1
                .IsVisibleInLegend = False
            End With
            If SerialPort1.IsOpen Then
                SerialPort1.Close()
            Else
                SerialPort1.BaudRate = ComboBox20.SelectedItem
                SerialPort1.PortName = ComboBox19.SelectedItem
                SerialPort1.Open()
                Dim ka As String = CStr(SerialPort1.ReadLine.ToString)
                If Not ka = "" And Not ka = " " Then

                    If IsNumeric(CDbl(ka)) = True Then
                        If Not CDbl(ka) > CDbl(TextBox16.Text) And Not CStr(ka) = " " And Not CStr(ka) = "" And Not CDbl(ka) < 0 And Not CStr(ka) = "" And Not CDbl(ka) < CDbl(TextBox18.Text) Then
                            ListView8.Items.Add(ListView8.Items.Count + 1)
                            ListView8.Items(i - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                            ListView8.Items(i - 1).SubItems.Add(CDbl(ka))
                            Chart3.Series(0).Points.AddXY((ListView8.Items(i - 1).SubItems(1).Text.ToString), CDbl(ListView8.Items(i - 1).SubItems(2).Text))
                            Chart3.Series(0).Points(i - 2).MarkerStyle = MarkerStyle.None
                            Chart3.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                            Chart3.Series(0).Points(i - 1).MarkerSize = 10
                            Chart3.Series(0).Points(i - 1).MarkerColor = Color.Red
                            Chart3.Series(0).Points(i - 1).LabelForeColor = Color.Blue
                            Chart3.Series(0).Points(i - 2).Label = ""
                            Chart3.Series(0).Points(i - 1).Label = CDbl(ListView8.Items(i - 1).SubItems(1).Text) & ";" & CDbl(ListView8.Items(i - 1).SubItems(2).Text)
                            If CheckBox5.Checked = True Then
                                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                            Else
                                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = TextBox20.Text + TextBox23.Text
                                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = TextBox19.Text - TextBox23.Text
                                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                            End If
                            'mean

                            If Not TextBox14.Text = "" Then
                                TextBox14.Text = Format(CDbl((TextBox14.Text * (ListView8.Items.Count - 1) + (CDec(ListView8.Items(i - 1).SubItems(2).Text))) / ListView8.Items.Count), "0.00E00")
                            Else
                                TextBox14.Text = Format(CDbl(ListView8.Items(i - 1).SubItems(2).Text), "0.000E00")
                            End If

                            'standard deviation

                            If Not TextBox13.Text = "" Then
                                TextBox13.Text = Format(CDbl(Math.Sqrt(((TextBox13.Text ^ 2) * (ListView8.Items.Count - 2) / (ListView8.Items.Count - 1)) + ((Math.Abs(CDec(ListView8.Items(i - 1).SubItems(2).Text) - CDec(TextBox14.Text))) ^ 2 / (ListView8.Items.Count - 1)))), "0.00E00")
                            Else
                                TextBox13.Text = 0
                            End If
                        Else

                            'Jika Data error dan memilih menggunakan data sebelumnya
                            If CheckBox3.Checked And i >= 2 Then
                                ListView8.Items.Add(ListView8.Items.Count + 1)
                                ListView8.Items(i - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                                ListView8.Items(i - 1).SubItems.Add(CDbl(ListView8.Items(i - 2).SubItems(2).Text))
                                Chart3.Series(0).Points.AddXY((ListView8.Items(i - 1).SubItems(1).Text.ToString), CDbl(ListView8.Items(i - 1).SubItems(2).Text))
                                Chart3.Series(0).Points(i - 2).MarkerStyle = MarkerStyle.None
                                Chart3.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                                Chart3.Series(0).Points(i - 1).MarkerSize = 10
                                Chart3.Series(0).Points(i - 1).MarkerColor = Color.Red
                                Chart3.Series(0).Points(i - 1).LabelForeColor = Color.Blue
                                Chart3.Series(0).Points(i - 2).Label = ""
                                Chart3.Series(0).Points(i - 1).Label = CDbl(ListView8.Items(i - 1).SubItems(1).Text) & ";" & CDbl(ListView8.Items(i - 1).SubItems(2).Text)
                                If CheckBox5.Checked = True Then
                                    Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
                                    Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
                                    Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                                    Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                                Else
                                    Chart3.ChartAreas("ChartArea1").AxisY.Maximum = TextBox20.Text + TextBox23.Text
                                    Chart3.ChartAreas("ChartArea1").AxisY.Minimum = TextBox19.Text - TextBox23.Text
                                    Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                                    Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                                End If
                                'mean

                                If Not TextBox14.Text = "" Then
                                    TextBox14.Text = Format(CDbl((TextBox14.Text * (ListView8.Items.Count - 1) + (CDec(ListView8.Items(i - 1).SubItems(2).Text))) / ListView8.Items.Count), "0.00E00")
                                Else
                                    TextBox14.Text = Format(CDbl(ListView8.Items(i - 1).SubItems(2).Text), "0.000E00")
                                End If

                                'standard deviation

                                If Not TextBox13.Text = "" Then
                                    TextBox13.Text = Format(CDbl(Math.Sqrt(((TextBox13.Text ^ 2) * (ListView8.Items.Count - 2) / (ListView8.Items.Count - 1)) + ((Math.Abs(CDec(ListView8.Items(i - 1).SubItems(2).Text) - CDec(TextBox14.Text))) ^ 2 / (ListView8.Items.Count - 1)))), "0.00E00")
                                Else
                                    TextBox13.Text = 0
                                End If
                            Else
                                wait(1)
                            End If
                            TextBox22.Text = TextBox22.Text + 1
                        End If
                    Else
                        If CheckBox3.Checked And i >= 2 Then
                            ListView8.Items.Add(ListView8.Items.Count + 1)
                            ListView8.Items(i - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                            ListView8.Items(i - 1).SubItems.Add(CDbl(ListView8.Items(i - 2).SubItems(2).Text))
                            Chart3.Series(0).Points.AddXY((ListView8.Items(i - 1).SubItems(1).Text.ToString), CDbl(ListView8.Items(i - 1).SubItems(2).Text))
                            Chart3.Series(0).Points(i - 2).MarkerStyle = MarkerStyle.None
                            Chart3.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                            Chart3.Series(0).Points(i - 1).MarkerSize = 10
                            Chart3.Series(0).Points(i - 1).MarkerColor = Color.Red
                            Chart3.Series(0).Points(i - 1).LabelForeColor = Color.Blue
                            Chart3.Series(0).Points(i - 2).Label = ""
                            Chart3.Series(0).Points(i - 1).Label = CDbl(ListView8.Items(i - 1).SubItems(1).Text) & ";" & CDbl(ListView8.Items(i - 1).SubItems(2).Text)
                            If CheckBox5.Checked = True Then
                                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                            Else
                                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = TextBox20.Text + TextBox23.Text
                                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = TextBox19.Text - TextBox23.Text
                                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                            End If
                            'mean

                            If Not TextBox14.Text = "" Then
                                TextBox14.Text = Format(CDbl((TextBox14.Text * (ListView8.Items.Count - 1) + (CDec(ListView8.Items(i - 1).SubItems(2).Text))) / ListView8.Items.Count), "0.00E00")
                            Else
                                TextBox14.Text = Format(CDbl(ListView8.Items(i - 1).SubItems(2).Text), "0.000E00")
                            End If

                            'standard deviation

                            If Not TextBox13.Text = "" Then
                                TextBox13.Text = Format(CDbl(Math.Sqrt(((TextBox13.Text ^ 2) * (ListView8.Items.Count - 2) / (ListView8.Items.Count - 1)) + ((Math.Abs(CDec(ListView8.Items(i - 1).SubItems(2).Text) - CDec(TextBox14.Text))) ^ 2 / (ListView8.Items.Count - 1)))), "0.00E00")
                            Else
                                TextBox13.Text = 0
                            End If
                        Else
                            wait(1)
                        End If
                        'Jika data bukan numeric
                        TextBox22.Text = TextBox22.Text + 1
                    End If
                Else 'Jika data tidak terbaca
                    If CheckBox3.Checked And i >= 2 Then
                        ListView8.Items.Add(ListView8.Items.Count + 1)
                        ListView8.Items(i - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                        ListView8.Items(i - 1).SubItems.Add(CDbl(ListView8.Items(i - 2).SubItems(2).Text))
                        Chart3.Series(0).Points.AddXY((ListView8.Items(i - 1).SubItems(1).Text.ToString), CDbl(ListView8.Items(i - 1).SubItems(2).Text))
                        Chart3.Series(0).Points(i - 2).MarkerStyle = MarkerStyle.None
                        Chart3.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
                        Chart3.Series(0).Points(i - 1).MarkerSize = 10
                        Chart3.Series(0).Points(i - 1).MarkerColor = Color.Red
                        Chart3.Series(0).Points(i - 1).LabelForeColor = Color.Blue
                        Chart3.Series(0).Points(i - 2).Label = ""
                        Chart3.Series(0).Points(i - 1).Label = CDbl(ListView8.Items(i - 1).SubItems(1).Text) & ";" & CDbl(ListView8.Items(i - 1).SubItems(2).Text)
                        If CheckBox5.Checked = True Then
                            Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
                            Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
                            Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                            Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                        Else
                            Chart3.ChartAreas("ChartArea1").AxisY.Maximum = TextBox20.Text + TextBox23.Text
                            Chart3.ChartAreas("ChartArea1").AxisY.Minimum = TextBox19.Text - TextBox23.Text
                            Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                            Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                        End If
                        'mean

                        If Not TextBox14.Text = "" Then
                            TextBox14.Text = Format(CDbl((TextBox14.Text * (ListView8.Items.Count - 1) + (CDec(ListView8.Items(i - 1).SubItems(2).Text))) / ListView8.Items.Count), "0.00E00")
                        Else
                            TextBox14.Text = Format(CDbl(ListView8.Items(i - 1).SubItems(2).Text), "0.000E00")
                        End If

                        'standard deviation

                        If Not TextBox13.Text = "" Then
                            TextBox13.Text = Format(CDbl(Math.Sqrt(((TextBox13.Text ^ 2) * (ListView8.Items.Count - 2) / (ListView8.Items.Count - 1)) + ((Math.Abs(CDec(ListView8.Items(i - 1).SubItems(2).Text) - CDec(TextBox14.Text))) ^ 2 / (ListView8.Items.Count - 1)))), "0.00E00")
                        Else
                            TextBox13.Text = 0
                        End If
                    Else
                        wait(1)
                    End If
                    TextBox22.Text = TextBox22.Text + 1
                End If
                ' TextBox22.Text = TextBox22.Text + 1
            End If
            'by time
            If RadioButton1.Checked = True Then
                If TextBox21.Text = TextBox10.Text Then
                    Timer6.Stop()
                    Timer4.Stop()
                    MsgBox("Calibration Finish")
                End If
            End If
            'by value
            If RadioButton3.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    Timer6.Stop()
                    Timer4.Stop()
                    MsgBox("Calibration Finish")
                End If
            End If
            'by value+time
            If RadioButton2.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    ti = ti + 1
                    If ti = TextBox10.Text Then
                        Timer6.Stop()
                        Timer4.Stop()
                        MsgBox("Calibration Finish")
                    End If
                Else
                    ti = 0
                End If
            End If

            'by stable value desired
            If RadioButton5.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    ts = ts + 1
                    If ts = TextBox24.Text Then
                        Timer6.Stop()
                        Timer4.Stop()
                        MsgBox("Calibration Finish")
                    End If
                Else
                    ts = 0
                End If
            End If

            'by stable value desired using int time
            If RadioButton4.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    ' Timer5.Interval = TextBox17.Text
                    'Timer5.Start()
                    Label46.Text = CDec(Label46.Text) + 1
                    ck = CDec(ck + CDec(ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text))
                    ' Label47.Text = "Second"
                    ' Label45.Text = "Calibration will finish in :"

                    If Label46.Text = CDec(TextBox24.Text) And CStr(CDec(CDec(TextBox8.Text) - CDec(TextBox11.Text)) * CDec(Label46.Text)) < (CDec(TextBox8.Text - CDec(ck))) < (CDec(CDec(TextBox8.Text) + CDec(TextBox11.Text)) * CDec(Label46.Text)) Then
                        Timer6.Stop()
                        Timer4.Stop()
                        MsgBox("Calibration Finish")
                    End If
                Else
                    Label46.Text = "0"
                    'Label45.Text = "Instable data detected"
                    'Label47.Text = ""
                End If
            End If
            'BY UNDETECTED STABLE VALUE 
            If RadioButton6.Checked = True Then
                Label46.Text = "NaN"
                ListView9.Items.Add(ListView9.Items.Count + 1)
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(CDbl(ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text))
                'substarct
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(CDbl(TextBox13.Text))
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(CDbl(ListView10.Items(ListView10.Items.Count - 1).SubItems(1).Text) - ListView9.Items(ListView9.Items.Count - 1).SubItems(2).Text)

                ListView9.Items(ListView9.Items.Count - 1).EnsureVisible()
                If ListView9.Items.Count > 1 Then
                    If Math.Abs(CDec(ListView9.Items(ListView9.Items.Count - 1).SubItems(3).Text)) <= CDec(ListView10.Items(ListView10.Items.Count - 1).SubItems(2).Text) Then
                        ListView9.Items(ListView9.Items.Count - 1).BackColor = Color.Red
                        ListView9.Items(ListView9.Items.Count - 1).ForeColor = Color.White

                        ListView10.Items(ListView10.Items.Count - 1).SubItems(3).Text = Format(CDbl(ListView9.Items(ListView9.Items.Count - 1).SubItems(2).Text), "00.00E00")
                        ListView10.Items(ListView10.Items.Count - 1).SubItems(4).Text = CDec(ListView9.Items(ListView9.Items.Count - 1).SubItems(4).Text)
                        'wait(1)

                        Timer6.Stop()
                        Timer4.Stop()
                        Timer5.Stop()
                        Timer3.Stop()
                        Timer2.Stop()
                        Timer1.Stop()
                        MsgBox("Calibration Finish")
                    End If
                End If
            End If

        Catch exp As Exception
            TextBox22.Text = TextBox22.Text + 1
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            TabControl1.SelectedIndex = 2
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub


    Private Sub Button77_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button77.Click
        Try
            CmbScanPort.Items.Clear()
            Dim myPort As Array
            Dim i As Integer
            myPort = IO.Ports.SerialPort.GetPortNames
            CmbScanPort.Items.AddRange(myPort)
            i = CmbScanPort.Items.Count
            i = i - i
            Try
                CmbScanPort.SelectedIndex = i
            Catch ex As Exception
                Dim result As DialogResult
                result = MessageBox.Show("Com Port not detected", "Warning!!!", MessageBoxButtons.OK)
                CmbScanPort.Text = ""
                CmbScanPort.Items.Clear()
                Call Form1_Load(Me, e)
            End Try
            Button28.Enabled = True
            CmbScanPort.DroppedDown = True
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button72_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button72.Click
        Try
            Dim SaveFile As New SaveFileDialog
            SaveFile.FileName = ""
            SaveFile.Filter = "Text Files (*.txt)|*.txt"
            SaveFile.Title = "Save"
            SaveFile.ShowDialog()
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)
            Dim col As ColumnHeader
            Dim columnnames As String = ""
            For Each col In ListView8.Columns
                If String.IsNullOrEmpty(columnnames) Then
                    columnnames = col.Text
                Else
                    columnnames &= "|" & col.Text
                End If
            Next
            Write.Write(columnnames & vbCrLf)
            For Me.baris = 1 To ListView8.Items.Count - 1
                Write.Write(ListView8.Items(baris - 1).SubItems(0).Text & "|" & ListView8.Items(baris - 1).SubItems(1).Text & "|" & ListView8.Items(baris - 1).SubItems(2).Text & vbCrLf)
            Next baris
            Write.Close()
        Catch d As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button73_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button73.Click
        Try
            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|SVG (*.svg)|*.svg|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
            saveFileDialog1.FilterIndex = 2
            saveFileDialog1.RestoreDirectory = True
            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                Dim format As ChartImageFormat = ChartImageFormat.Bmp
                If saveFileDialog1.FileName.EndsWith("bmp") Then
                    format = ChartImageFormat.Bmp
                Else
                    If saveFileDialog1.FileName.EndsWith("jpg") Then
                        format = ChartImageFormat.Jpeg
                    Else
                        If saveFileDialog1.FileName.EndsWith("emf") Then
                            format = ChartImageFormat.Emf
                        Else
                            If saveFileDialog1.FileName.EndsWith("gif") Then
                                format = ChartImageFormat.Gif
                            Else
                                If saveFileDialog1.FileName.EndsWith("png") Then
                                    format = ChartImageFormat.Png
                                Else
                                    If saveFileDialog1.FileName.EndsWith("tif") Then
                                        format = ChartImageFormat.Tiff
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                Chart3.SaveImage(saveFileDialog1.FileName, format)
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button74_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button74.Click
        Try
            Chart3.Series(0).Points.Clear()
            ListView8.Items.Clear()
            TextBox13.Text = ""
            TextBox14.Text = ""
            TextBox21.Text = 0
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button75.Click
        Chart3.Series(0).Points.Clear()

        wait(1)

        For i = 1 To ListView8.Items.Count
            Chart3.Series(0).Points.AddXY((ListView8.Items(i - 1).SubItems(1).Text.ToString), CDbl(ListView8.Items(i - 1).SubItems(2).Text))
            If i > 1 Then
                Chart3.Series(0).Points(i - 2).MarkerStyle = MarkerStyle.None
                Chart3.Series(0).Points(i - 2).Label = ""
            End If
            Chart3.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
            Chart3.Series(0).Points(i - 1).MarkerSize = 10
            Chart3.Series(0).Points(i - 1).MarkerColor = Color.Red
            Chart3.Series(0).Points(i - 1).LabelForeColor = Color.Blue

            Chart3.Series(0).Points(i - 1).Label = (ListView8.Items(i - 1).SubItems(1).Text) & " ; " & CDec(ListView8.Items(i - 1).SubItems(2).Text)
            If CheckBox5.Checked = True Then
                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
            Else
                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Val(TextBox20.Text) + TextBox23.Text
                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Val(TextBox19.Text) - TextBox23.Text
                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
            End If
            wait(0.01)
        Next
    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox16.SelectedIndexChanged
        Label8.Text = ComboBox16.Text
        Label9.Text = ComboBox16.Text
        Label25.Text = ComboBox16.Text
        Label15.Text = ComboBox16.Text
        Label18.Text = ComboBox16.Text

    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox18.SelectedIndexChanged
        Chart3.Titles.Clear()
        Try
            If ComboBox18.Text = "Custom" Then
                Form11.Show()
            Else
                ListView8.Columns(2).Text = ComboBox18.Text.ToString

                'TextBox49.Text = "Time VS " + ComboBox3.Text
                Chart3.Titles.Add("Time VS " + ComboBox18.Text)
                'Chart2.Series(0).Name = ComboBox3.Text
                Chart3.ChartAreas(0).AxisX.Title = "Time(hh:mm:ss)"
                Chart3.ChartAreas(0).AxisX.Enabled = AxisEnabled.True
                Chart3.ChartAreas(0).AxisY.Enabled = AxisEnabled.True
                Chart3.ChartAreas(0).AxisY.Title = ComboBox18.Text
              
            End If

        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        Label8.Text = ComboBox16.Text & TextBox12.Text
        Label9.Text = ComboBox16.Text & TextBox12.Text
        Label25.Text = ComboBox16.Text & TextBox12.Text
        Label15.Text = ComboBox16.Text & TextBox12.Text
        Label18.Text = ComboBox16.Text & TextBox12.Text
        Chart3.ChartAreas(0).AxisY.Title = ComboBox18.Text & "(" & ComboBox16.Text & TextBox12.Text & ")"
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        TextBox8.Enabled = True
        TextBox10.Enabled = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        TextBox8.Enabled = True
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        TextBox8.Enabled = True
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        TextBox8.Enabled = True
        Label46.Text = 0
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        TextBox20.Enabled = False
        TextBox19.Enabled = False
        TextBox23.Enabled = False
        If CheckBox5.Checked = False Then
            TextBox20.Enabled = True
            TextBox19.Enabled = True
            TextBox23.Enabled = True
            CheckBox6.Checked = True
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        TextBox20.Enabled = True
        TextBox19.Enabled = True
        TextBox23.Enabled = True
        If CheckBox6.Checked = False Then
            TextBox20.Enabled = False
            TextBox19.Enabled = False
            TextBox23.Enabled = False
            CheckBox5.Checked = True
        End If
    End Sub

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged
        If TextBox25.Text < TextBox17.Text Then
            TextBox25.Text = ""
            MsgBox("Interval must be bigger than Gate Time")
        End If
    End Sub

    Private Sub Button78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button78.Click
        Timer6.Enabled = True
        Timer6.Interval = TextBox17.Text * 1000
        Timer6.Start()
        '  Timer5.Enabled = True
        ' Timer5.Interval = TextBox17.Text * 1000
        'Timer5.Start()
    End Sub

    Private Sub Timer6_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer6.Tick
        Try
            TextBox21.Text = Val(TextBox21.Text) + 1
            Dim i As Integer = ListView8.Items.Count + 1

            With TrendLine
                .ChartType = SeriesChartType.Line
                .Color = Color.DodgerBlue
                .BorderWidth = 1
                .IsVisibleInLegend = False
            End With

            Dim ka As String = 7 * (Math.Exp(-(ListView8.Items.Count + 1)))

            ListView8.Items.Add(ListView8.Items.Count + 1)
            ListView8.Items(i - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
            ' If CheckBox3.Checked = True Then
            'ListView8.Items.Clear()
            'i = 2
            'Else

            ListView8.Items(i - 1).SubItems.Add(CDbl(ka))
            ListView8.Items(i - 1).EnsureVisible()
            'End If
            Chart3.Series(0).Points.AddXY((ListView8.Items(i - 1).SubItems(1).Text.ToString), CDbl(ListView8.Items(i - 1).SubItems(2).Text))
            If i > 1 Then
                Chart3.Series(0).Points(i - 2).MarkerStyle = MarkerStyle.None
                Chart3.Series(0).Points(i - 2).Label = ""
                ListView8.Items(i - 1).BackColor = Color.Blue
                ListView8.Items(i - 1).ForeColor = Color.White
                ListView8.Items(i - 2).BackColor = Color.White
                ListView8.Items(i - 2).ForeColor = Color.Black
            End If

            Chart3.Series(0).Points(i - 1).MarkerStyle = MarkerStyle.Circle
            Chart3.Series(0).Points(i - 1).MarkerSize = 10
            Chart3.Series(0).Points(i - 1).MarkerColor = Color.Red
            Chart3.Series(0).Points(i - 1).LabelForeColor = Color.Blue

            Chart3.Series(0).Points(i - 1).Label = (ListView8.Items(i - 1).SubItems(1).Text) & ";" & CDec(ListView8.Items(i - 1).SubItems(2).Text)
            If CheckBox5.Checked = True Then
                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
            Else
                Chart3.ChartAreas("ChartArea1").AxisY.Maximum = Val(TextBox20.Text) + TextBox23.Text
                Chart3.ChartAreas("ChartArea1").AxisY.Minimum = Val(TextBox19.Text) - TextBox23.Text
                Chart3.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                Chart3.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
            End If
            'mean

            If Not TextBox14.Text = "" Then
                TextBox14.Text = Format(CDbl((TextBox14.Text * (ListView8.Items.Count - 1) + (CDec(ListView8.Items(i - 1).SubItems(2).Text))) / ListView8.Items.Count), "0.00E00")
            Else
                TextBox14.Text = (CDec(ListView8.Items(i - 1).SubItems(2).Text))
            End If

            'standard deviation
            If ListView8.Items.Count > 1 Then
                If Not TextBox13.Text = "" Then
                    TextBox13.Text = Format(CDbl(Math.Sqrt(((TextBox13.Text ^ 2) * (ListView8.Items.Count - 2) / (ListView8.Items.Count - 1)) + ((Math.Abs(CDec(ListView8.Items(i - 1).SubItems(2).Text) - CDec(TextBox14.Text))) ^ 2 / (ListView8.Items.Count - 1)))), "0.00E00")
                Else
                    TextBox13.Text = 0
                End If
            Else
                TextBox13.Text = 0
            End If

            ' TextBox22.Text = TextBox22.Text + 1

            'by time
            If RadioButton1.Checked = True Then
                If TextBox21.Text = TextBox10.Text Then
                    Timer6.Stop()
                    Timer4.Stop()
                    MsgBox("Calibration Finish")
                End If
            End If
            'by value
            If RadioButton3.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    Timer6.Stop()
                    Timer4.Stop()
                    MsgBox("Calibration Finish")
                End If
            End If
            'by value+time
            If RadioButton2.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    ti = ti + 1
                    If ti = TextBox10.Text Then
                        Timer6.Stop()
                        Timer4.Stop()
                        MsgBox("Calibration Finish")
                    End If
                Else
                    ti = 0
                End If
            End If

            'by stable value desired
            If RadioButton5.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    ts = ts + 1
                    If ts = TextBox24.Text Then
                        Timer6.Stop()
                        Timer4.Stop()
                        MsgBox("Calibration Finish")
                    End If
                Else
                    ts = 0
                End If
            End If

            'by stable value desired using int time
            If RadioButton4.Checked = True Then
                If ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text <= TextBox8.Text + TextBox11.Text And ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text >= TextBox8.Text - TextBox11.Text Then
                    ' Timer5.Interval = TextBox17.Text
                    'Timer5.Start()
                    Label46.Text = CDec(Label46.Text) + 1
                    ck = CDec(ck + CDec(ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text))
                    ' Label47.Text = "Second"
                    ' Label45.Text = "Calibration will finish in :"

                    If Label46.Text = CDec(TextBox24.Text) And CStr(CDec(CDec(TextBox8.Text) - CDec(TextBox11.Text)) * CDec(Label46.Text)) < (CDec(TextBox8.Text - CDec(ck))) < (CDec(CDec(TextBox8.Text) + CDec(TextBox11.Text)) * CDec(Label46.Text)) Then
                        Timer6.Stop()
                        Timer4.Stop()
                        MsgBox("Calibration Finish")
                    End If
                Else
                    Label46.Text = "0"
                    'Label45.Text = "Instable data detected"
                    'Label47.Text = ""
                End If
            End If
            'BY UNDETECTED STABLE VALUE 
            If RadioButton6.Checked = True Then
                Label46.Text = "NaN"
                ListView9.Items.Add(ListView9.Items.Count + 1)
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(Date.Now.ToString("HH:mm:ss"))
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(CDbl(ListView8.Items(ListView8.Items.Count - 1).SubItems(2).Text))
                'substarct
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(CDbl(TextBox13.Text))
                ListView9.Items(ListView9.Items.Count - 1).SubItems.Add(CDbl(ListView10.Items(ListView10.Items.Count - 1).SubItems(1).Text) - ListView9.Items(ListView9.Items.Count - 1).SubItems(2).Text)

                ListView9.Items(ListView9.Items.Count - 1).EnsureVisible()
                If ListView9.Items.Count > 1 Then
                    If Math.Abs(CDec(ListView9.Items(ListView9.Items.Count - 1).SubItems(3).Text)) <= CDec(ListView10.Items(ListView10.Items.Count - 1).SubItems(2).Text) Then
                        ListView9.Items(ListView9.Items.Count - 1).BackColor = Color.Red
                        ListView9.Items(ListView9.Items.Count - 1).ForeColor = Color.White

                        ListView10.Items(ListView10.Items.Count - 1).SubItems(3).Text = Format(CDbl(ListView9.Items(ListView9.Items.Count - 1).SubItems(2).Text), "00.00E00")
                        ListView10.Items(ListView10.Items.Count - 1).SubItems(4).Text = CDec(ListView9.Items(ListView9.Items.Count - 1).SubItems(4).Text)
                        'wait(1)

                        Timer6.Stop()
                        Timer4.Stop()
                        Timer5.Stop()
                        Timer3.Stop()
                        Timer2.Stop()
                        Timer1.Stop()
                        MsgBox("Calibration Finish")
                    End If
                End If
            End If


        Catch exp As Exception
            'TextBox22.Text = TextBox22.Text + 1
        End Try

    End Sub
    Private Sub Button79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button79.Click
        Try
            Dim SaveFile As New SaveFileDialog
            SaveFile.FileName = ""
            SaveFile.Filter = "Text Files (*.txt)|*.txt"
            SaveFile.Title = "Save"
            SaveFile.ShowDialog()
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)
            Dim col As ColumnHeader
            Dim columnnames As String = ""
            For Each col In ListView10.Columns
                If String.IsNullOrEmpty(columnnames) Then
                    columnnames = col.Text
                Else
                    columnnames &= "|" & col.Text
                End If
            Next
            Write.Write(columnnames & vbCrLf)
            For Me.baris = 1 To ListView10.Items.Count - 1
                Write.Write(ListView10.Items(baris - 1).SubItems(0).Text & "|" & ListView10.Items(baris - 1).SubItems(1).Text & "|" & ListView10.Items(baris - 1).SubItems(2).Text & vbCrLf)
            Next baris
            Write.Close()
        Catch d As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button80.Click
        ListView10.Items.Clear()
    End Sub

    Private Sub Button82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button82.Click
        Try
            Dim SaveFile As New SaveFileDialog
            SaveFile.FileName = ""
            SaveFile.Filter = "Text Files (*.txt)|*.txt"
            SaveFile.Title = "Save"
            SaveFile.ShowDialog()
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)
            Dim col As ColumnHeader
            Dim columnnames As String = ""
            ' For Each col In ListView3.Columns
            'If String.IsNullOrEmpty(columnnames) Then
            'columnnames = col.Text
            'Else
            ' columnnames &= "|" & col.Text
            ' End If
            ' Next
            ' Write.Write(columnnames & vbCrLf)
            ' For Me.baris = 1 To ListView3.Items.Count - 1
            Write.Write(Label14.Text & " = " & TextBox14.Text & vbCrLf)
            Write.Write(Label12.Text & " = " & TextBox13.Text & vbCrLf)
            Write.Write(Label37.Text & " = " & TextBox21.Text & vbCrLf)
            Write.Write(Label39.Text & " = " & TextBox22.Text & vbCrLf)
            Write.Close()
        Catch d As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button81_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button81.Click
        TextBox14.Text = ""
        TextBox13.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        TextBox8.Enabled = True
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        Label46.Text = TextBox10.Text() - 1
    End Sub

    Private Sub Button85_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button85.Click
        ListView10.Items.Add(ListView10.Items.Count + 1)
        ListView10.Items(ListView10.Items.Count - 1).SubItems.Add(TextBox8.Text)
        ListView10.Items(ListView10.Items.Count - 1).SubItems.Add(TextBox15.Text)
        ListView10.Items(ListView10.Items.Count - 1).SubItems.Add("")
        ListView10.Items(ListView10.Items.Count - 1).SubItems.Add("")
    End Sub

    Private Sub Button84_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button84.Click
        Try
            Dim SaveFile As New SaveFileDialog
            SaveFile.FileName = ""
            SaveFile.Filter = "Text Files (*.txt)|*.txt"
            SaveFile.Title = "Save"
            SaveFile.ShowDialog()
            Dim Write As New System.IO.StreamWriter(SaveFile.FileName)
            Dim col As ColumnHeader
            Dim columnnames As String = ""
            For Each col In ListView9.Columns
                If String.IsNullOrEmpty(columnnames) Then
                    columnnames = col.Text
                Else
                    columnnames &= "|" & col.Text
                End If
            Next
            Write.Write(columnnames & vbCrLf)
            For Me.baris = 1 To ListView9.Items.Count - 1
                Write.Write(ListView9.Items(baris - 1).SubItems(0).Text & "|" & ListView9.Items(baris - 1).SubItems(1).Text & "|" & ListView9.Items(baris - 1).SubItems(2).Text & vbCrLf)
            Next baris
            Write.Close()
        Catch d As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button86_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button86.Click
        TextBox8.Text = ""
        TextBox11.Text = ""
    End Sub

    Private Sub Button83_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button83.Click
        ListView9.Items.Clear()
    End Sub

    Private Sub Button76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button76.Click
        Form12.Show()

    End Sub

End Class
