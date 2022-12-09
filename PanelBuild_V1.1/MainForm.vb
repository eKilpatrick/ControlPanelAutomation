Imports System.Drawing
Imports System.IO.Ports
Public Class MainForm
    Public serialNum As String = ""
    Public drawingNum As String = ""
    Public ActivePN As String = ""

    Public sRobotStatus As Boolean = True
    Public gRobotStatus As Boolean = True

    Public screwbot As New RobotConnect("screwbot", "169.254.245.125", "169.254.245.130", 55555)
    Public gripper As New RobotConnect("gripbot", "169.254.245.127", "169.254.245.130", 55557)

    Public spComRobot As SerialPort
    Public bportOpen As Boolean = False

    Public UIParts As New List(Of Item)

#Region "Form Controls"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ssnTxtBox.Focus()
        ActivePartComplete()
        CheckStorage()
        Dim RecentOrder = GetRecentOrder()
        If RecentOrder IsNot Nothing Then
            'ssnTxtBox.Text = RecentOrder.item(0)
            'serialNum = RecentOrder.item(0)
            txtDN.Text = RecentOrder.item(1)
            drawingNum = RecentOrder.item(1)
        End If
        QueueTimer.Start()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        screwbot.Connect()
        gripper.Connect()
    End Sub

    Private Sub Form1_Close(sender As Object, e As EventArgs) Handles MyBase.Closing
        screwbot.RTDE_Client.Disconnect(False)
        gripper.RTDE_Client.Disconnect(False)
    End Sub
    Private Sub SimScrewbot_CheckedChanged(sender As Object, e As EventArgs) Handles SimScrewbot.CheckedChanged
        If SimScrewbot.Checked Then
            sRobotStatus = False
        Else
            sRobotStatus = True
        End If
    End Sub

    Private Sub SimGripbot_CheckedChanged(sender As Object, e As EventArgs) Handles SimGripbot.CheckedChanged
        If SimGripbot.Checked Then
            gRobotStatus = False
        Else
            gRobotStatus = True
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles QueueTimer.Tick
        QueueTimer.Stop()
        CabQueue()
        QueueTimer.Start()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        'When the active tab is Home vs not Home
        If TabControl1.SelectedTab IsNot TabControl1.TabPages.Item(0) Then
            LblActiveDN.Text = "DN: " & txtDN.Text
        ElseIf TabControl1.SelectedTab Is TabControl1.TabPages.Item(0) Then
            LblActiveDN.Text = ""
        End If
        'When the active tab is Robot Control vs not Robot control
        If TabControl1.SelectedTab Is TabControl1.TabPages.Item(2) Then
            RobotTimer.Start()
        Else
            RobotTimer.Stop()
        End If
        'When the active tab is UI vs not UI
        If TabControl1.SelectedTab Is TabControl1.TabPages.Item(3) Then
            BtnDrawUI_Click(sender, e)
            UITimer.Start()
        Else
            UITimer.Stop()
        End If
        'When the active tab is Log vs no Log
        If TabControl1.SelectedTab Is TabControl1.TabPages.Item(4) Then
            BtnSaveLog.Visible = True
        Else
            BtnSaveLog.Visible = False
        End If
        'When the active tab is storage vs no storage
        If TabControl1.SelectedTab Is TabStorage Then
            StorageTimer.Start()
        Else
            StorageTimer.Stop()
        End If

    End Sub

    Private Sub BtnSaveLog_Click(sender As Object, e As EventArgs) Handles BtnSaveLog.Click
        Dim DebugFolder As String = String.Format("{0}\", Environment.CurrentDirectory) & "LOGS\"
        Dim DateNow As String = DateTime.Now.ToString("yy" & "-" & "MM" & "-" & "dd" & "_" & "HHmmss")
        Using sw As New System.IO.StreamWriter(DebugFolder & DateNow & ".txt")
            sw.Write(txtLog.Text)
        End Using
    End Sub

    Private Sub RobotTimer_Tick(sender As Object, e As EventArgs) Handles RobotTimer.Tick
        RobotTimer.Stop()
#Region "Screwbot Calc"
        Dim tx = screwbot.UrOutputs.target_TCP_pose(0) * 1000
        Dim ty = screwbot.UrOutputs.target_TCP_pose(1) * 1000
        Dim tz = screwbot.UrOutputs.target_TCP_pose(2) * 1000
        Dim trx = screwbot.UrOutputs.target_TCP_pose(3)
        Dim try1 = screwbot.UrOutputs.target_TCP_pose(4)
        Dim trz = screwbot.UrOutputs.target_TCP_pose(5)

        lbltx.Text = "X: " & Math.Round(tx, 3)
        lblty.Text = "Y: " & Math.Round(ty, 3)
        lbltz.Text = "Z: " & Math.Round(tx, 3)
        lbltrx.Text = "RX: " & Math.Round(trx, 3)
        lbltry.Text = "RX: " & Math.Round(try1, 3)
        lbltrz.Text = "RZ: " & Math.Round(trz, 3)

        Dim ax = screwbot.UrOutputs.actual_TCP_pose(0) * 1000
        Dim ay = screwbot.UrOutputs.actual_TCP_pose(1) * 1000
        Dim az = screwbot.UrOutputs.actual_TCP_pose(2) * 1000
        Dim arx = screwbot.UrOutputs.actual_TCP_pose(3)
        Dim ary = screwbot.UrOutputs.actual_TCP_pose(4)
        Dim arz = screwbot.UrOutputs.actual_TCP_pose(5)

        lblax.Text = "X: " & Math.Round(ax, 3)
        lblay.Text = "Y: " & Math.Round(ay, 3)
        lblaz.Text = "Z: " & Math.Round(az, 3)
        lblarx.Text = "RX: " & Math.Round(arx, 3)
        lblary.Text = "RY: " & Math.Round(ary, 3)
        lblarz.Text = "RZ: " & Math.Round(arz, 3)

        Dim ex As Double = Math.Round(((ax - tx) / ax * 100), 3)
        Dim ey As Double = Math.Round(((ay - ty) / ay * 100), 3)
        Dim ez As Double = Math.Round(((az - tz) / az * 100), 3)
        Dim erx As Double = Math.Round(((arx - trx) / arx * 100), 3)
        Dim ery As Double = Math.Round(((ary - try1) / ary * 100), 3)
        Dim erz As Double = Math.Round(((arz - trz) / arz * 100), 3)

        lblex.Text = "X: " & ex & "%"
        lbley.Text = "Y: " & ey & "%"
        lblez.Text = "Z: " & ez & "%"
        lblerx.Text = "RX: " & erx & "%"
        lblery.Text = "RY: " & ery & "%"
        lblerz.Text = "RZ: " & erz & "%"

        TxtSpeedScrewbot.Text = screwbot.UrOutputs.target_speed_fraction * 100
#End Region

#Region "Gripper Calc"
        Dim gtx = gripper.UrOutputs.target_TCP_pose(0) * 1000
        Dim gty = gripper.UrOutputs.target_TCP_pose(1) * 1000
        Dim gtz = gripper.UrOutputs.target_TCP_pose(2) * 1000
        Dim gtrx = gripper.UrOutputs.target_TCP_pose(3)
        Dim gtry = gripper.UrOutputs.target_TCP_pose(4)
        Dim gtrz = gripper.UrOutputs.target_TCP_pose(5)

        lblgtx.Text = "X: " & Math.Round(gtx, 3)
        lblgty.Text = "Y: " & Math.Round(gty, 3)
        lblgtz.Text = "Z: " & Math.Round(gtx, 3)
        lblgtrx.Text = "RX: " & Math.Round(gtrx, 3)
        lblgtry.Text = "RX: " & Math.Round(gtry, 3)
        lblgtrz.Text = "RZ: " & Math.Round(gtrz, 3)

        Dim gax = gripper.UrOutputs.actual_TCP_pose(0) * 1000
        Dim gay = gripper.UrOutputs.actual_TCP_pose(1) * 1000
        Dim gaz = gripper.UrOutputs.actual_TCP_pose(2) * 1000
        Dim garx = gripper.UrOutputs.actual_TCP_pose(3)
        Dim gary = gripper.UrOutputs.actual_TCP_pose(4)
        Dim garz = gripper.UrOutputs.actual_TCP_pose(5)

        lblgax.Text = "X: " & Math.Round(gax, 3)
        lblgay.Text = "Y: " & Math.Round(gay, 3)
        lblgaz.Text = "Z: " & Math.Round(gaz, 3)
        lblgarx.Text = "RX: " & Math.Round(garx, 3)
        lblgary.Text = "RY: " & Math.Round(gary, 3)
        lblgarz.Text = "RZ: " & Math.Round(garz, 3)

        Dim gex = Math.Round(((gax - gtx) / gax * 100), 3)
        Dim gey = Math.Round(((gay - gty) / gay * 100), 3)
        Dim gez = Math.Round(((gaz - gtz) / gaz * 100), 3)
        Dim gerx = Math.Round(((garx - gtrx) / garx * 100), 3)
        Dim gery = Math.Round(((gary - gtry) / gary * 100), 3)
        Dim gerz = Math.Round(((garz - gtrz) / garz * 100), 3)

        lblgex.Text = "X: " & gex & "%"
        lblgey.Text = "Y: " & gey & "%"
        lblgez.Text = "Z: " & gez & "%"
        lblgerx.Text = "RX: " & gerx & "%"
        lblgery.Text = "RY: " & gery & "%"
        lblgerz.Text = "RZ: " & gerz & "%"

        TxtSpeedGripper.Text = gripper.UrOutputs.target_speed_fraction * 100
#End Region
        RobotTimer.Start()
    End Sub

    Private Sub btnSMode_Click(sender As Object, e As EventArgs) Handles btnSMode.Click
        If screwbot.CheckRemoteMode() Then
            lblSMode.Text = "Mode: Remote"
            lblSMode.ForeColor = Color.Green
        Else
            lblSMode.Text = "Mode: Local"
            lblSMode.ForeColor = Color.Red
        End If
    End Sub

    Private Sub btnGMode_Click(sender As Object, e As EventArgs) Handles btnGMode.Click
        If gripper.CheckRemoteMode() Then
            lblGMode.Text = "Mode: Remote"
            lblGMode.ForeColor = Color.Green
        Else
            lblGMode.Text = "Mode: Local"
            lblGMode.ForeColor = Color.Red
        End If
    End Sub

    Private Sub UITimer_Tick(sender As Object, e As EventArgs) Handles UITimer.Tick
        UITimer.Stop()
        BtnDrawUI_Click(sender, e)
        UITimer.Start()
    End Sub
    Private Sub BtnCloseRTDE_Click(sender As Object, e As EventArgs) Handles BtnCloseRTDE.Click
        screwbot.RTDE_Client.Ur_ControlPause()
    End Sub

    Private Sub BtnOpenRTDE_Click(sender As Object, e As EventArgs) Handles BtnOpenRTDE.Click
        screwbot.RTDE_Client.Ur_ControlStart()
    End Sub

    Private Sub BtnGOpenRTDE_Click(sender As Object, e As EventArgs) Handles BtnGOpenRTDE.Click
        gripper.RTDE_Client.Ur_ControlStart()
    End Sub

    Private Sub BtnGCloseRTDE_Click(sender As Object, e As EventArgs) Handles BtnGCloseRTDE.Click
        gripper.RTDE_Client.Ur_ControlPause()
    End Sub

    Private Sub BtnDisconnectScrewbot_Click(sender As Object, e As EventArgs) Handles BtnDisconnectScrewbot.Click
        screwbot.RTDE_Client.Disconnect(True)
    End Sub

    Private Sub BtnDisconnectGripper_Click(sender As Object, e As EventArgs) Handles BtnDisconnectGripper.Click
        gripper.RTDE_Client.Disconnect(True)
    End Sub

    Private Sub PlusScrewSpeed_Click(sender As Object, e As EventArgs) Handles PlusScrewSpeed.Click
        Dim NewSpeed = screwbot.UrOutputs.target_speed_fraction + 0.05
        If NewSpeed > 1 Or NewSpeed < 0 Then
            MsgBox("That speed is not attainable")
        Else
            screwbot.UrInputs.speed_slider_fraction = NewSpeed
            screwbot.RTDE_Client.Send_Ur_Inputs()
        End If
    End Sub

    Private Sub MinusScrewSpeed_Click(sender As Object, e As EventArgs) Handles MinusScrewSpeed.Click
        Dim NewSpeed = screwbot.UrOutputs.target_speed_fraction - 0.05
        If NewSpeed > 1 Or NewSpeed < 0 Then
            MsgBox("That speed is not attainable")
        Else
            screwbot.UrInputs.speed_slider_fraction = NewSpeed
            screwbot.RTDE_Client.Send_Ur_Inputs()
        End If
    End Sub

    Private Sub PlusGripSpeed_Click(sender As Object, e As EventArgs) Handles PlusGripSpeed.Click
        Dim NewSpeed = gripper.UrOutputs.target_speed_fraction + 0.05
        If NewSpeed > 1 Or NewSpeed < 0 Then
            MsgBox("That speed is not attainable")
        Else
            gripper.UrInputs.speed_slider_fraction = NewSpeed
            gripper.RTDE_Client.Send_Ur_Inputs()
        End If
    End Sub

    Private Sub MinusGripSpeed_Click(sender As Object, e As EventArgs) Handles MinusGripSpeed.Click
        Dim NewSpeed = gripper.UrOutputs.target_speed_fraction - 0.05
        If NewSpeed > 1 Or NewSpeed < 0 Then
            MsgBox("That speed is not attainable")
        Else
            gripper.UrInputs.speed_slider_fraction = NewSpeed
            gripper.RTDE_Client.Send_Ur_Inputs()
        End If
    End Sub
    Private Sub StorageTimer_Tick(sender As Object, e As EventArgs) Handles StorageTimer.Tick
        StorageTimer.Stop()

        CheckStorage()

        StorageTimer.Start()
    End Sub

    Private Sub ssnTxtBox_TextChanged(sender As Object, e As EventArgs) Handles ssnTxtBox.TextChanged
        If ssnTxtBox.Text.Trim = "75002531-1" Then
            serialNum = ssnTxtBox.Text.Trim()

            TabControl1.SelectedTab = TabUI
            Me.Update()
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub
#End Region

#Region "Run Order"
    Private Sub runOrderBtn_Click(sender As Object, e As EventArgs) Handles runOrderBtn.Click
        Try
            serialNum = ssnTxtBox.Text.Trim()
            If serialNum <> "Serial Number" Then
                BackgroundWorker1.RunWorkerAsync()
                TabControl1.SelectedTab = TabUI
                Me.Update()
            Else
                MsgBox("Please enter a serial number")
            End If
        Catch
            MsgBox("Already running. Restart program if needed.")
            Return
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        CheckSerialNumber(serialNum, drawingNum)
        WriteActiveOrder(serialNum, drawingNum)
        ActivePartComplete() 'Added this in case the program gets stopped during the middle of running a part so it resets before the next part runs.
        Dim sPanel As String = GetPanel(drawingNum)
        Dim bCheckCount = CheckCount(drawingNum)

        CheckForIllegalCrossThreadCalls = False
        While bCheckCount = True
            RunTopPart(drawingNum, sPanel)
            bCheckCount = CheckCount(drawingNum)
        End While
    End Sub

    Private Sub RunSinglePart_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles RunSinglePart.DoWork
        CheckForIllegalCrossThreadCalls = False
        SetIncomplete(drawingNum, ActivePN)
        RunPart(ActivePN, drawingNum)
    End Sub

    Public Sub RunTopPart(drawingNum As String, sPanel As String)
        Dim sPN As String = GetFirstItem(drawingNum, sPanel)

        If sPN <> Nothing Then
            RunPart(sPN, drawingNum)
        Else
            MsgBox("No more parts found")
        End If
    End Sub

    Public Sub RunPart(sPN As String, drawingNum As String)
        Dim splitPN As String

        If sPN <> "" And sPN.Contains("_") Then
            splitPN = sPN.Split("_")(0)
        ElseIf sPN <> "" Then
            splitPN = sPN.Split("-")(0)
        Else
            splitPN = Nothing
        End If

        PartProgressBar.Value = 0
        LblProgress.Text = "Progress: Grabbing part " & splitPN

        Dim Part As New Part()
        'Gets all of the information from the database and intializes all of the part's vars
        Part.InitializeVariables(sPN, splitPN, drawingNum, screwbot)
        'Starts the log for the part that was just initialized
        UpdateLog("BeginPart", Part)
        'Sends all of the part's vars to the robots' registers through the RTDE client

        If screwbot.RTDE_Client.ControlConnected = False Or gripper.RTDE_Client.ControlConnected = False Then
            If screwbot.RTDE_Client.ControlConnected = False Then
                MsgBox("Screwbot RTDE Control not started correctly")
            End If
            If gripper.RTDE_Client.ControlConnected = False Then
                MsgBox("Gripper RTDE Control not started correctly")
            End If
            Application.Restart()
        End If

        Part.SendVarsRobot(screwbot, gripper)

        'Starts the gripper's process of grabbing the part from storage and placing on the panel
        If gRobotStatus Then
            ResetOutputVars(gripper)
            GetPart(gripper)
            ChangeStorage(Part.BinNum, -1)
        End If

        'Starts the screwbot's process of getting the screw from the presentor
        If sRobotStatus Then
            ResetOutputVars(screwbot)
            If Part.CommonName <> "KNIFESWITCH" Then
                ScrewPart(Part, 2, screwbot, spComRobot, bportOpen)
            Else
                ScrewPart(Part, 1, screwbot, spComRobot, bportOpen)
            End If
            LightGateCheck(Part, screwbot, spComRobot, bportOpen)
        End If

        'Waits for the part to be placed on the panel
        If gRobotStatus Then
            PlacedOnPanel(gripper)

            'If gripper doesn't need to wait for the first screw, then wait till it gets home before screwing.
            If Part.WaitToLeave = False Then
                GoingBackHome(gripper)
            End If

            'Updates the log for the theoretical/actual positions at the storage bin and at the panel
            UpdateLog("Storage", Part)
            UpdateLog("PlacePanel", Part)

            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Part Placed"
        Else
            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Part Placed (Simulation)"
        End If

        'Begins screwing the part and waits until complete
        If sRobotStatus Then
            AllowScrewing(Part, screwbot)

            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Started Screwing 1"

            DoneScrewing(screwbot)

            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Finished Screwing 1"

            'Updates the log for the theoretical/actual positions at the screw pres and at the panel hole
            UpdateLog("ScrewPresentor", Part)
            If Part.WhichScrew Then
                UpdateLog("ScrewPanel2", Part)
            Else
                UpdateLog("ScrewPanel1", Part)
            End If


            'Waits for the screwbot to go back home if the gripper is activated so they don't collide if the gripper waited on the panel with the part
            If gRobotStatus Then
                BackHome(screwbot)
            End If
        Else
            PartProgressBar.PerformStep()
            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Finished Screwing 1 (Simulation)"
        End If

        'The gripper can now leave the panel, given it isn't already at home for certain parts
        If gRobotStatus Then
            LeavePanel(gripper)
        End If

        'Starts the screwbot's process of getting the 2nd screw from the presentor and screws part until complete
        If sRobotStatus Then
            ResetOutputVars(screwbot)
            If Part.CommonName <> "KNIFESWITCH" Then
                ScrewPart(Part, 1, screwbot, spComRobot, bportOpen)
            Else
AskKnifeswitch:
                Dim response = MsgBox("This part is a knifeswitch and requires opening the latch prior to the 2nd screw", vbYesNo)
                If response = vbNo Then
                    GoTo AskKnifeswitch
                End If
                ScrewPart(Part, 2, screwbot, spComRobot, bportOpen)
            End If
            LightGateCheck(Part, screwbot, spComRobot, bportOpen)

            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Started Screwing 2"

            DoneScrewing(screwbot)
            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Finished Screwing 2"

            'Updates the log for the theoretical/actual positions at the 2nd Hole screwed (changes for some parts whether it's screw1 or 2
            If Part.WhichScrew Then
                UpdateLog("ScrewPanel2", Part)
            Else
                UpdateLog("ScrewPanel1", Part)
            End If

        Else
            PartProgressBar.PerformStep()
            PartProgressBar.PerformStep()
            LblProgress.Text = "Progress: Finished Screwing 2 (Simulation)"
        End If

        'Updates the part to complete in Y_Panel_Build_Positions
        SendComplete(sPN, drawingNum)

        'Sets the part to inactive in Y_Panel_Build_Robot_Offsets
        ActivePartComplete()
    End Sub

    'Need to add bit to change Screw1 and Screw2 if TB. +0.015 in Y for 2 and -0.015 in Y for 1
    Public Sub UpdateLog(cmd As String, ByVal Part As Part)
        If cmd = "Storage" Then
            txtLog.Text &= "Storage Theoretical: {" & Part.StorageBinCoord(0) & "," & Part.StorageBinCoord(1) & "," & Part.StorageBinCoord(2) & "," & Part.StorageBinCoord(3) & "," & Part.StorageBinCoord(4) & "," & Part.StorageBinCoord(5) & "}" & Environment.NewLine
            txtLog.Text &= "Storage Actual: {" & gripper.UrOutputs.output_double_register_24 & "," & gripper.UrOutputs.output_double_register_25 & "," & gripper.UrOutputs.output_double_register_26 & "," & gripper.UrOutputs.output_double_register_27 & "," & gripper.UrOutputs.output_double_register_28 & "," & gripper.UrOutputs.output_double_register_29 & "}" & Environment.NewLine

        ElseIf cmd = "PlacePanel" Then
            txtLog.Text &= "Place Panel Theoretical: {" & Part.PlacePanel(0) & "," & Part.PlacePanel(1) & "," & Part.PlacePanel(2) & "," & Part.PlacePanel(3) & "," & Part.PlacePanel(4) & "," & Part.PlacePanel(5) & "}" & Environment.NewLine
            txtLog.Text &= "Place Panel Actual: {" & gripper.UrOutputs.output_double_register_30 & "," & gripper.UrOutputs.output_double_register_31 & "," & gripper.UrOutputs.output_double_register_32 & "," & gripper.UrOutputs.output_double_register_33 & "," & gripper.UrOutputs.output_double_register_34 & "," & gripper.UrOutputs.output_double_register_35 & "}" & Environment.NewLine

        ElseIf cmd = "ScrewPresentor" Then
            txtLog.Text &= "Screw Pres Theoretical: {" & Part.ScrewPresLoc(0) & "," & Part.ScrewPresLoc(1) & "," & Part.ScrewPresLoc(2) & "," & Part.ScrewPresLoc(3) & "," & Part.ScrewPresLoc(4) & "," & Part.ScrewPresLoc(5) & "}" & Environment.NewLine
            txtLog.Text &= "Screw Pres Actual: {" & screwbot.UrOutputs.output_double_register_24 & "," & screwbot.UrOutputs.output_double_register_25 & "," & screwbot.UrOutputs.output_double_register_26 & "," & screwbot.UrOutputs.output_double_register_27 & "," & screwbot.UrOutputs.output_double_register_28 & "," & screwbot.UrOutputs.output_double_register_29 & "}" & Environment.NewLine

        ElseIf cmd = "ScrewPanel1" Then
            If Part.CommonName = "12PT TB" Then
                Part.Screw1PosAbovePanel(1) -= 0.015
            End If
            txtLog.Text &= "Screw Panel 1 Theoretical: {" & Part.Screw1PosAbovePanel(0) & "," & Part.Screw1PosAbovePanel(1) & "," & Part.Screw1PosAbovePanel(2) & "," & Part.Screw1PosAbovePanel(3) & "," & Part.Screw1PosAbovePanel(4) & "," & Part.Screw1PosAbovePanel(5) & "}" & Environment.NewLine
            txtLog.Text &= "Screw Panel 1 Actual: {" & screwbot.UrOutputs.output_double_register_30 & "," & screwbot.UrOutputs.output_double_register_31 & "," & screwbot.UrOutputs.output_double_register_32 & "," & screwbot.UrOutputs.output_double_register_33 & "," & screwbot.UrOutputs.output_double_register_34 & "," & screwbot.UrOutputs.output_double_register_35 & "}" & Environment.NewLine

        ElseIf cmd = "ScrewPanel2" Then
            If Part.CommonName = "12PT TB" Then
                Part.Screw2PosAbovePanel(1) += 0.015
            End If
            txtLog.Text &= "Screw Panel 2 Theoretical: {" & Part.Screw2PosAbovePanel(0) & "," & Part.Screw2PosAbovePanel(1) & "," & Part.Screw2PosAbovePanel(2) & "," & Part.Screw2PosAbovePanel(3) & "," & Part.Screw2PosAbovePanel(4) & "," & Part.Screw2PosAbovePanel(5) & "}" & Environment.NewLine
            txtLog.Text &= "Screw Panel 2 Actual: {" & screwbot.UrOutputs.output_double_register_36 & "," & screwbot.UrOutputs.output_double_register_37 & "," & screwbot.UrOutputs.output_double_register_38 & "," & screwbot.UrOutputs.output_double_register_39 & "," & screwbot.UrOutputs.output_double_register_40 & "," & screwbot.UrOutputs.output_double_register_41 & "}" & Environment.NewLine

        ElseIf cmd = "BeginPart" Then
            txtLog.Text &= "-----------------------------------" & Environment.NewLine
            txtLog.Text &= Part.CommonName & ", " & Part.sPN & Environment.NewLine
            txtLog.Text &= "-----------------------------------" & Environment.NewLine

        Else
            Console.WriteLine("Invalid cmd to update log")
        End If
    End Sub
#End Region

#Region "UI"
    Public Structure Item
        Public PartNum As String
        Public Rect As Rectangle
        Public color As Color

        Public BuildPosX1 As Double
        Public BuildPosY1 As Double
        Public BuildPosZ1 As Double
        Public BuildPosX2 As Double
        Public BuildPosY2 As Double
        Public BuildPosZ2 As Double
    End Structure

    Private Sub BtnDrawUI_Click(sender As Object, e As EventArgs) Handles BtnDrawUI.Click
        UIParts.Clear()

        Using g As Graphics = PanelDWG.CreateGraphics()
            Dim BuildPositions As DataTable = GetBuildForUI(drawingNum)

            For Each row As DataRow In BuildPositions.Rows
                If row.Item(0) <> "72480649001:1" Then
                    Dim xCoord As Integer = row.Item(1) * 15
                    Dim yCoord As Integer = 442 - (row.Item(5) * 15)
                    Dim width As Integer = Math.Abs(row.Item(1) - row.Item(4)) * 15
                    Dim length As Integer = Math.Abs(row.Item(2) - row.Item(5)) * 15

                    Dim Rect As New Rectangle(xCoord, yCoord, width, length)

                    Dim blackPen As New Pen(Brushes.Black)

                    'Adds parts on the panel to a list of the structure Item for reference when clicking them.
                    Dim part As New Item
                    part.PartNum = row.Item(0)
                    part.Rect = Rect
                    If row.Item(7) = 0 Then
                        part.color = Color.Red
                    Else
                        part.color = Color.Green
                    End If
                    part.BuildPosX1 = row.Item(1)
                    part.BuildPosY1 = row.Item(2)
                    part.BuildPosZ1 = row.Item(3)
                    part.BuildPosX2 = row.Item(4)
                    part.BuildPosY2 = row.Item(5)
                    part.BuildPosZ2 = row.Item(6)
                    UIParts.Add(part)

                    g.DrawRectangle(blackPen, Rect)
                    If row.Item(7) = 0 Then
                        g.FillRectangle(Brushes.Red, Rect)
                    Else
                        g.FillRectangle(Brushes.Green, Rect)
                    End If
                    Invalidate()
                End If
            Next
        End Using
    End Sub

    Private Sub SurfaceSelection_Click(sender As Object, e As MouseEventArgs) Handles PanelDWG.MouseClick
        UITimer.Stop()
        For Each item As Item In UIParts
            If item.Rect.X < e.X And e.X < (item.Rect.X + item.Rect.Width) And item.Rect.Y < e.Y And e.Y < (item.Rect.Y + item.Rect.Height) Then
                Dim response = MsgBox("Would you like to run this part(" & item.PartNum & ")?", vbYesNo)
                If response = vbYes Then
                    ActivePN = item.PartNum
                    RunSinglePart.RunWorkerAsync()
                End If
            End If
        Next
        UITimer.Start()
    End Sub
#End Region

#Region "Storage"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = CInt(TextBox1.Text) + 1
        ChangeStorage(8, 1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = CInt(TextBox1.Text) - 1
        ChangeStorage(8, -1)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox6.Text = CInt(TextBox6.Text) + 1
        ChangeStorage(9, 1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox6.Text = CInt(TextBox6.Text) - 1
        ChangeStorage(9, -1)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox9.Text = CInt(TextBox9.Text) + 1
        ChangeStorage(10, 1)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox9.Text = CInt(TextBox9.Text) - 1
        ChangeStorage(10, -1)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox12.Text = CInt(TextBox12.Text) + 1
        ChangeStorage(11, 1)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox12.Text = CInt(TextBox12.Text) - 1
        ChangeStorage(11, -1)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox15.Text = CInt(TextBox15.Text) + 1
        ChangeStorage(12, 1)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox15.Text = CInt(TextBox15.Text) - 1
        ChangeStorage(12, -1)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        TextBox18.Text = CInt(TextBox18.Text) + 1
        ChangeStorage(13, 1)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox18.Text = CInt(TextBox18.Text) - 1
        ChangeStorage(13, -1)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox21.Text = CInt(TextBox21.Text) + 1
        ChangeStorage(14, 1)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox21.Text = CInt(TextBox21.Text) - 1
        ChangeStorage(14, -1)
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        TextBox42.Text = CInt(TextBox42.Text) + 1
        ChangeStorage(15, 1)
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        TextBox42.Text = CInt(TextBox42.Text) - 1
        ChangeStorage(15, -1)
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        TextBox39.Text = CInt(TextBox39.Text) + 1
        ChangeStorage(16, 1)
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        TextBox39.Text = CInt(TextBox39.Text) - 1
        ChangeStorage(16, -1)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        TextBox36.Text = CInt(TextBox36.Text) + 1
        ChangeStorage(17, 1)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        TextBox36.Text = CInt(TextBox36.Text) - 1
        ChangeStorage(17, -1)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        TextBox33.Text = CInt(TextBox33.Text) + 1
        ChangeStorage(18, 1)
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        TextBox33.Text = CInt(TextBox33.Text) - 1
        ChangeStorage(18, -1)
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        TextBox30.Text = CInt(TextBox30.Text) + 1
        ChangeStorage(19, 1)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        TextBox30.Text = CInt(TextBox30.Text) - 1
        ChangeStorage(19, -1)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        TextBox27.Text = CInt(TextBox27.Text) + 1
        ChangeStorage(20, 1)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox27.Text = CInt(TextBox27.Text) - 1
        ChangeStorage(20, -1)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        TextBox24.Text = CInt(TextBox24.Text) + 1
        ChangeStorage(21, 1)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        TextBox24.Text = CInt(TextBox24.Text) - 1
        ChangeStorage(21, -1)
    End Sub

    Private Sub RefreshStorage_Click_1(sender As Object, e As EventArgs) Handles RefreshStorage.Click
        CheckStorage()
    End Sub
#End Region

End Class
