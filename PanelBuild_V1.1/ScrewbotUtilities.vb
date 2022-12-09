Imports System.IO.Ports
Module ScrewbotUtilities
    Public Sub SetMountOrigin(ByRef Part As Part)
        If Part.PHoleX < 0.05555 Then 'Distance from Top left mounting screw to the side panel stop. HARDCODED (probably won't change)**********
            Part.arrSBot(0) -= Part.PHoleX
        Else
            Part.arrSBot(0) = -0.04675 'HardCoded X value for the side stop. only used if the panel hits the side HARDCODED *******************************
        End If

        Part.arrSBot(1) -= Part.PHoleY
    End Sub

    Public Sub ScrewPositionCalibration(ByRef Part As Part)
        If Part.PartRotate = False Then
            'Hole closest to the bottom left
            Part.HoleX1 = CDbl(Part.arrSBot(0)) + Part.X1 + Part.HX1
            Part.HoleY1 = CDbl(Part.arrSBot(1)) + Part.Y1 + Part.HY1

            'Hole furthest from the bottom left
            Part.HoleX2 = CDbl(Part.arrSBot(0)) + Part.X1 + Part.HX2
            Part.HoleY2 = CDbl(Part.arrSBot(1)) + Part.Y1 + Part.HY2
        Else
            'Hole closest to the bottom left
            Part.HoleX1 = CDbl(Part.arrSBot(0)) + Part.X1 + Part.HY1
            Part.HoleY1 = CDbl(Part.arrSBot(1)) + Part.Y2 - Part.HX1

            'Hole furthest from the bottom left
            Part.HoleX2 = CDbl(Part.arrSBot(0)) + Part.X1 + Part.HY2
            Part.HoleY2 = CDbl(Part.arrSBot(1)) + Part.Y2 - Part.HX2
        End If

        'This section is for hardcoded offsets, should NOT be here permanently
        If Part.CommonName = "CIRCUIT BREAKER" Then
            Part.HoleX1 += 0.0121 - 0.001
            Part.HoleX2 += 0.0121 - 0.001

            Part.HoleY1 += 0.0005
            Part.HoleY2 += 0.0005
        ElseIf Part.CommonName = "FUSEBLOCK" Then
            Part.HoleX1 -= 0.0041
            Part.HoleX2 -= 0.0041

            Part.HoleY1 -= 0.0004
            Part.HoleY2 -= 0.0004
        ElseIf Part.CommonName = "KNIFESWITCH" Then
            Part.HoleX1 -= 0.0016
            Part.HoleX2 -= 0.0016

            Part.HoleY1 += 0.0005
            Part.HoleY2 += 0.0005
        ElseIf Part.CommonName = "12PT TB" Then
            Part.HoleX1 -= 0.0025
            Part.HoleX2 -= 0.0025

            'Part.HoleY1 -= 0.001
            'Part.HoleY2 -= 0.001
        End If
    End Sub

    Public Sub SetActiveScrewPresentor(ByRef Part As Part, ByRef screwbot As RobotConnect)
        'Presentor 1
        If Part.ScrewType = 0 Then
            Dim PresLoc As DataRow = GetOrigin("PRESENTOR1")
            Part.ScrewPresLoc = {PresLoc(1), PresLoc(2), PresLoc(3), PresLoc(4), PresLoc(5), PresLoc(6)}
            Part.ScrewRemoval = True
            Part.WhichPres = 0
            'Presentor 2
        ElseIf Part.ScrewType = 1 Then
            Dim PresLoc As DataRow = GetOrigin("PRESENTOR2")
            Part.ScrewPresLoc = {PresLoc(1), PresLoc(2), PresLoc(3), PresLoc(4), PresLoc(5), PresLoc(6)}
            Part.ScrewRemoval = True
            Part.WhichPres = 1
            'Presentor 3
        ElseIf Part.ScrewType = 2 Then
            Part.ScrewPresLoc = {0, 0, 0, 0, 0, 0}
            Part.ScrewRemoval = False
        Else
            MsgBox("SCREW TYPE NOT KNOWN/NOT AVAILABLE")
        End If
    End Sub

    Public Sub TBScrewFix(ByRef Part As Part, ByRef screwbot As RobotConnect)
        If Part.CommonName = "12PT TB" Then
            Part.tbOffset = True
        End If
    End Sub

    'Plays the program NewScrewPart1.urp or NewScrewPart2.urp
    Public Sub ScrewPart(ByRef Part As Part, Screw As Integer, ByRef screwbot As RobotConnect, spComRobot As SerialPort, bportOpen As Boolean)
        If Screw = 1 Then
            Part.WhichScrew = False
            screwbot.UrInputs.input_bit_register_67 = Part.WhichScrew
            If screwbot.RTDE_Client.Send_Ur_Inputs() Then
                Console.WriteLine("Sent successfully to robot")
            Else
                Console.WriteLine("Error sending info to robot")
            End If
            screwbot.StartProgram("/programs/_NewScrewPart.urp")
        ElseIf Screw = 2 Then
            Part.WhichScrew = True
            screwbot.UrInputs.input_bit_register_67 = Part.WhichScrew
            If screwbot.RTDE_Client.Send_Ur_Inputs() Then
                Console.WriteLine("Sent successfully to robot")
            Else
                Console.WriteLine("Error sending info to robot")
            End If
            screwbot.StartProgram("/programs/_NewScrewPart.urp")
        Else
            MsgBox("Part has more than 3 screws???")
            Exit Sub
        End If
        ProgramStarted(screwbot)
    End Sub

    Public Sub AllowScrewing(ByRef Part As Part, ByRef screwbot As RobotConnect)
        Part.WaitToScrew = False
        screwbot.UrInputs.input_bit_register_66 = Part.WaitToScrew
        If screwbot.RTDE_Client.Send_Ur_Inputs() Then
            Console.WriteLine("Successfully updated screwbot registers")
        Else
            Console.WriteLine("Error updating screwbot registers")
        End If
    End Sub

    'False means the incorrect outcome occurred (i.e. screw when there shouldn't be one or vice versa
    Public Sub LightGateCheck(ByRef Part As Part, ByRef screwbot As RobotConnect, ByRef spComRobot As SerialPort, ByRef bportOpen As Boolean)
        'Add LightGate code here
        Dim bLightGate As Boolean

        AtLightgate1(screwbot)

        'commented ReceiveSerialData out for testing while the lightgate isn't working
        bLightGate = False 'ReceiveSerialData(spComRobot, bportOpen)
        'At this point there should NOT be a screw
        If bLightGate = True Then
            'This means there is a screw BAD
            MsgBox("There should not be a screw now, manual inspection required")
            screwbot.UrInputs.input_bit_register_69 = True
            screwbot.UrInputs.input_bit_register_70 = True
            screwbot.UrInputs.input_bit_register_71 = False
            screwbot.RTDE_Client.Send_Ur_Inputs()
            Application.Restart()
        Else
            'This means there is not a screw GOOD
            screwbot.UrInputs.input_bit_register_69 = False
            screwbot.UrInputs.input_bit_register_70 = True
            screwbot.UrInputs.input_bit_register_71 = False
            screwbot.RTDE_Client.Send_Ur_Inputs()
        End If

        'Parts that don't need to be screwed don't go to the lightgate twice and require no screw
        If Part.ScrewRemoval = False Then
            Exit Sub
        End If
LightGate2:
        AtLightGate2(screwbot)
        'At this point there SHOULD be a screw
        'commented ReceiveSerialData out for testing while the lightgate isn't working
        bLightGate = True 'ReceiveSerialData(spComRobot, bportOpen)
        If bLightGate = True Then
            'This means there is a screw GOOD
            screwbot.UrInputs.input_bit_register_69 = True
            screwbot.UrInputs.input_bit_register_70 = False
            screwbot.UrInputs.input_bit_register_71 = True
            screwbot.RTDE_Client.Send_Ur_Inputs()
        Else
            'This means there is not a screw BAD
            screwbot.UrInputs.input_bit_register_69 = False
            screwbot.UrInputs.input_bit_register_70 = False
            screwbot.UrInputs.input_bit_register_71 = True
            screwbot.RTDE_Client.Send_Ur_Inputs()
            WaitToReAttemptScrew(screwbot)
            GoTo LightGate2
        End If

    End Sub
#Region "Screwbot Status Updates"
    Public Sub ProgramStarted(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_64 <> True
            'Wait until program has started
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub AtLightgate1(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_65 <> True
            'Wait until the screwbot is at the lightgate
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub AtPresentor(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_66 <> True
            'Wait until the screwbot is at the presentor
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub AtLightGate2(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_67 <> True
            'Wait until the screwbot is at the lightgate
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub Screwing(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_68 <> True
            'Wait until the screwbot has started screwing
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub DoneScrewing(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_69 <> True
            'Wait until the Screwbot is done screwing
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub BackHome(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_70 <> True
            'Wait until the screwbot is back home
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub ProgramEnded(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_71 <> True
            'Wait until the program is complete
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub WaitToReAttemptScrew(ByRef screwbot As RobotConnect)
        While screwbot.UrOutputs.output_bit_register_72 <> True
            'Wait until the screwbot is ready to reattempt to get a screw
            Threading.Thread.Sleep(250)
        End While
    End Sub
#End Region
End Module
