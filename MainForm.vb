Imports DAL1
Public Class MainForm

    Dim screwbot As New RobotConnect("screwbot", "169.254.245.127", "169.254.245.130", 55555)
    Dim gripbot As New RobotConnect("gripbot", "169.254.245.125", "169.254.245.140", 55557)
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Top = 0
        Me.Left = 0
        Me.Height = Screen.PrimaryScreen.WorkingArea.Height
        Me.Width = Screen.PrimaryScreen.WorkingArea.Width
        QueueTimer.Start()
        txtSN.Text = DatabaseQueries.RecentOrder().split(",")(0)
        txtDN.Text = DatabaseQueries.RecentOrder().split(",")(1)
        DatabaseQueries.CheckStorage()
    End Sub

    Private Sub RunOrderBtn_Click(sender As Object, e As EventArgs) Handles RunOrderBtn.Click
        RunOrderThread.RunWorkerAsync()
    End Sub

    Private Sub RunOrderThread_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles RunOrderThread.DoWork
        Try
            Dim SN = txtSN.Text
            Dim DN = txtDN.Text
            If SN Is Nothing Or DN Is Nothing Then
                MsgBox("You must provide both the serial number and drawing number of the order in order to run it")
                Exit Sub
            End If

            If DatabaseQueries.CheckExtraction(DN) Then
                DatabaseQueries.PastOrder(SN, DN)
                RunOrder.RunOrder(DN, screwbot, gripbot)
            Else
                MsgBox("The drawing number provided has yet to be extracted")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error beginning to run order: " & ex.Message)
        End Try
    End Sub

    Private Sub ExtractDrawingBtn_Click(sender As Object, e As EventArgs) Handles ExtractDrawingBtn.Click
        ExtractDrawingThread.RunWorkerAsync()
    End Sub

    Private Sub ExtractDrawingThread_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles ExtractDrawingThread.DoWork
        Try
            Dim DN = txtDN.Text
            If DN Is Nothing Then
                MsgBox("You must provide a drawing number to extract")
                Exit Sub
            End If
            If DatabaseQueries.CheckExtraction(DN) Then
                MsgBox("That drawing has already been extracted to the database")
            Else
                InventorExtraction.ExtractInventorData(DN)
                DatabaseQueries.FuseAssemblyExtract(DN)
            End If
        Catch ex As Exception
            MsgBox("Error during extraction: " & ex.Message)
        End Try
    End Sub

    Private Sub CloseBtn_Click(sender As Object, e As EventArgs) Handles CloseBtn.Click
        Dim proc = Process.GetProcessesByName("Panel_Build")
        For Each proc2 In proc
            proc2.Kill()
        Next
        Application.Exit()
    End Sub

    Private Sub MaximizeBtn_Click(sender As Object, e As EventArgs) Handles MaximizeBtn.Click
        If Me.Height = Screen.PrimaryScreen.WorkingArea.Height Then
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height / 2
            Me.Width = Screen.PrimaryScreen.WorkingArea.Width / 2
            Me.Top = Screen.PrimaryScreen.WorkingArea.Height / 4
            Me.Left = Screen.PrimaryScreen.WorkingArea.Width / 4
        Else
            Me.Top = 0
            Me.Left = 0
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height
            Me.Width = Screen.PrimaryScreen.WorkingArea.Width
        End If
    End Sub

    Private Sub MinimizeBtn_Click(sender As Object, e As EventArgs) Handles MinimizeBtn.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub QueueTimer_Tick(sender As Object, e As EventArgs) Handles QueueTimer.Tick
        QueueTimer.Stop()
        DatabaseQueries.UpdatePartsQueue()
        If screwbot.connected Then
            ScrewbotChecked.CheckState = CheckState.Checked
        Else
            ScrewbotChecked.CheckState = CheckState.Unchecked
        End If
        If gripbot.connected Then
            GripbotChecked.CheckState = CheckState.Checked
        Else
            GripbotChecked.CheckState = CheckState.Unchecked
        End If
        QueueTimer.Start()
    End Sub

    Private Sub ErrorResetBtn_Click(sender As Object, e As EventArgs) Handles ErrorResetBtn.Click
        'RunOrderThread.CancelAsync()
        'ExtractDrawingThread.CancelAsync()
        'screwbot.Stop()
        'gripbot.Stop()
        'screwbot.StartProgram("GoHome.urp")
        'gripbot.StartProgram("GoHome.urp")
        'Application.Restart()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = CInt(TextBox1.Text) + 1
        DatabaseQueries.ChangeStorage(8, 1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = CInt(TextBox1.Text) - 1
        DatabaseQueries.ChangeStorage(8, -1)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox6.Text = CInt(TextBox6.Text) + 1
        DatabaseQueries.ChangeStorage(9, 1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox6.Text = CInt(TextBox6.Text) - 1
        DatabaseQueries.ChangeStorage(9, -1)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox9.Text = CInt(TextBox9.Text) + 1
        DatabaseQueries.ChangeStorage(10, 1)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox9.Text = CInt(TextBox9.Text) - 1
        DatabaseQueries.ChangeStorage(10, -1)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox12.Text = CInt(TextBox12.Text) + 1
        DatabaseQueries.ChangeStorage(11, 1)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox12.Text = CInt(TextBox12.Text) - 1
        DatabaseQueries.ChangeStorage(11, -1)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox15.Text = CInt(TextBox15.Text) + 1
        DatabaseQueries.ChangeStorage(12, 1)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox15.Text = CInt(TextBox15.Text) - 1
        DatabaseQueries.ChangeStorage(12, -1)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        TextBox18.Text = CInt(TextBox18.Text) + 1
        DatabaseQueries.ChangeStorage(13, 1)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox18.Text = CInt(TextBox18.Text) - 1
        DatabaseQueries.ChangeStorage(13, -1)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox21.Text = CInt(TextBox21.Text) + 1
        DatabaseQueries.ChangeStorage(14, 1)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox21.Text = CInt(TextBox21.Text) - 1
        DatabaseQueries.ChangeStorage(14, -1)
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        TextBox42.Text = CInt(TextBox42.Text) + 1
        DatabaseQueries.ChangeStorage(15, 1)
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        TextBox42.Text = CInt(TextBox42.Text) - 1
        DatabaseQueries.ChangeStorage(15, -1)
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        TextBox39.Text = CInt(TextBox39.Text) + 1
        DatabaseQueries.ChangeStorage(16, 1)
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        TextBox39.Text = CInt(TextBox39.Text) - 1
        DatabaseQueries.ChangeStorage(16, -1)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        TextBox36.Text = CInt(TextBox36.Text) + 1
        DatabaseQueries.ChangeStorage(17, 1)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        TextBox36.Text = CInt(TextBox36.Text) - 1
        DatabaseQueries.ChangeStorage(17, -1)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        TextBox33.Text = CInt(TextBox33.Text) + 1
        DatabaseQueries.ChangeStorage(18, 1)
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        TextBox33.Text = CInt(TextBox33.Text) - 1
        DatabaseQueries.ChangeStorage(18, -1)
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        TextBox30.Text = CInt(TextBox30.Text) + 1
        DatabaseQueries.ChangeStorage(19, 1)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        TextBox30.Text = CInt(TextBox30.Text) - 1
        DatabaseQueries.ChangeStorage(19, -1)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        TextBox27.Text = CInt(TextBox27.Text) + 1
        DatabaseQueries.ChangeStorage(20, 1)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox27.Text = CInt(TextBox27.Text) - 1
        DatabaseQueries.ChangeStorage(20, -1)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        TextBox24.Text = CInt(TextBox24.Text) + 1
        DatabaseQueries.ChangeStorage(21, 1)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        TextBox24.Text = CInt(TextBox24.Text) - 1
        DatabaseQueries.ChangeStorage(21, -1)
    End Sub

    Private Sub RefreshStorage_Click(sender As Object, e As EventArgs) Handles RefreshStorage.Click
        DatabaseQueries.CheckStorage()
    End Sub


    Private Sub SimulateScrewbotCheck_CheckedChanged(sender As Object, e As EventArgs) Handles SimulateScrewbotCheck.CheckedChanged
        If SimulateScrewbotCheck.CheckState = CheckState.Checked Then
            screwbot.simulate = True
        Else
            screwbot.simulate = False
        End If
    End Sub

    Private Sub SimulateGripbotCheck_CheckedChanged(sender As Object, e As EventArgs) Handles SimulateGripbotCheck.CheckedChanged
        If SimulateGripbotCheck.CheckState = CheckState.Checked Then
            gripbot.simulate = True
        Else
            gripbot.simulate = False
        End If
    End Sub
End Class
