Imports System.IO.Ports
Imports System.Threading.Thread
Module ScrewProcess
    Public spComRobot As SerialPort
    Public bPortOpen As Boolean

    Public Sub StartScrewing(NextPart As NextPart, screwbot As RobotConnect, ksPt2 As Boolean)
        If NextPart.commonName = "KNIFESWITCH" And ksPt2 = False Then
            SecondScrew(NextPart, screwbot)
        ElseIf NextPart.commonName = "KNIFESWITCH" And ksPt2 = True Then
            FirstScrew(NextPart, screwbot)
        Else
            If NextPart.commonName = "12PT TB" Then
                NextPart.tbOffset = 0.015
            End If
            FirstScrew(NextPart, screwbot)
            If NextPart.commonName = "12PT TB" Then
                NextPart.tbOffset = -0.015
            End If
            SecondScrew(NextPart, screwbot)
        End If

    End Sub
    Public Sub FirstScrew(NextPart As NextPart, screwbot As RobotConnect)
        ScrewingProcess(NextPart, screwbot)
        If NextPart.PartRotate = True Then
            NextPart.ScrewPos1 = "(" & (NextPart.X1 + NextPart.hole2Y) & "," & (NextPart.Y1 + NextPart.hole2X + NextPart.tbOffset) & "," & "0.2" & ",2.2,2.2,0.0,0)"
            NextPart.ScrewPos2 = "(" & (NextPart.X1 + NextPart.hole2Y) & "," & (NextPart.Y1 + NextPart.hole2X + NextPart.tbOffset) & "," & (0.2 - NextPart.Offset) & ",2.2,2.2,0.0,0)"
            NextPart.ScrewPos3 = "(" & (NextPart.X1 + NextPart.hole2Y) & "," & (NextPart.Y1 + NextPart.hole2X) & "," & (0.2 - NextPart.Offset - NextPart.Depth) & ",2.2,2.2,0.0,0)"
        Else
            NextPart.ScrewPos1 = "(" & (NextPart.X1 + NextPart.hole2X) & "," & (NextPart.Y1 + NextPart.hole2Y + NextPart.tbOffset) & "," & "0.2" & ",0,-3.14,0.0,0)"
            NextPart.ScrewPos2 = "(" & (NextPart.X1 + NextPart.hole2X) & "," & (NextPart.Y1 + NextPart.hole2Y + NextPart.tbOffset) & "," & (0.2 - NextPart.Offset) & ",0,-3.14,0.0,0)"
            NextPart.ScrewPos3 = "(" & (NextPart.X1 + NextPart.hole2X) & "," & (NextPart.Y1 + NextPart.hole2Y) & "," & (0.2 - NextPart.Offset - NextPart.Depth) & ",0,-3.14,0.0,0)"
        End If

        screwbot.WriteTCP("panel", NextPart.ScrewPos1)
        screwbot.WriteTCP("godown", NextPart.ScrewPos2)
        If NextPart.commonName = "12PT TB" Then
            screwbot.WriteTCP("terminal", ">yes1<")
        Else
            screwbot.WriteTCP("terminal", ">no<")
        End If
        screwbot.WriteTCP("screw", NextPart.ScrewPos3)

        While screwbot.ReceiveTCP("done") = False
            'waiting for the screwbot to finish screwing the part
        End While
    End Sub

    Public Sub SecondScrew(NextPart As NextPart, screwbot As RobotConnect)
        ScrewingProcess(NextPart, screwbot)
        If NextPart.PartRotate = True Then
            NextPart.ScrewPos1 = "(" & (NextPart.X1 + NextPart.hole1Y) & "," & (NextPart.Y1 + NextPart.hole1X + NextPart.tbOffset) & "," & "0.2" & ",2.2,2.2,0.0,0)"
            NextPart.ScrewPos2 = "(" & (NextPart.X1 + NextPart.hole1Y) & "," & (NextPart.Y1 + NextPart.hole1X + NextPart.tbOffset) & "," & (0.2 - NextPart.Offset) & ",2.2,2.2,0.0,0)"
            NextPart.ScrewPos3 = "(" & (NextPart.X1 + NextPart.hole1Y) & "," & (NextPart.Y1 + NextPart.hole1X) & "," & (0.2 - NextPart.Offset - NextPart.Depth) & ",2.2,2.2,0.0,0)"
        Else
            NextPart.ScrewPos1 = "(" & (NextPart.X1 + NextPart.hole1X) & "," & (NextPart.Y1 + NextPart.hole1Y + NextPart.tbOffset) & "," & "0.2" & ",0,-3.14,0.0,0)"
            NextPart.ScrewPos2 = "(" & (NextPart.X1 + NextPart.hole1X) & "," & (NextPart.Y1 + NextPart.hole1Y + NextPart.tbOffset) & "," & (0.2 - NextPart.Offset) & ",0,-3.14,0.0,0)"
            NextPart.ScrewPos3 = "(" & (NextPart.X1 + NextPart.hole1X) & "," & (NextPart.Y1 + NextPart.hole1Y) & "," & (0.2 - NextPart.Offset - NextPart.Depth) & ",0,-3.14,0.0,0)"
        End If

        screwbot.WriteTCP("panel", NextPart.ScrewPos1)
        screwbot.WriteTCP("godown", NextPart.ScrewPos2)
        If NextPart.commonName = "12PT TB" Then
            screwbot.WriteTCP("terminal", ">yes2<")
        Else
            screwbot.WriteTCP("terminal", ">no<")
        End If
        screwbot.WriteTCP("screw", NextPart.ScrewPos3)

        While screwbot.ReceiveTCP("done") = False
            'waiting for the screwbot to finish screwing the part
        End While
    End Sub

    Public Sub ScrewingProcess(NextPart As NextPart, screwbot As RobotConnect)
        'Begin with the screw stuff....
        If NextPart.screwType = 0 Then
            NextPart.ScrewPresLoc = "(" & NextPart.Presentor1X & "," & NextPart.Presentor1Y & "," & NextPart.Presentor1Z & "," & NextPart.Presentor1RX & "," & NextPart.Presentor1RY & "," & NextPart.Presentor1RZ & ",0)" 'Location of bit to pick up Screw + 0.1 in the z axis
            Dim removeScrew1Z = NextPart.Presentor1Z + 0.1
            NextPart.ScrewPres2 = "(" & NextPart.Presentor1X & "," & NextPart.Presentor1Y & "," & removeScrew1Z & "," & NextPart.Presentor1RX & "," & NextPart.Presentor1RY & "," & NextPart.Presentor1RZ & ",0)" 'Location for the romoval of the Screw. This type of presenter removes vertically, so it is + 0.1 in the z axis
        ElseIf NextPart.screwType = 1 Then
            NextPart.ScrewPresLoc = "(" & NextPart.Presentor2X & "," & NextPart.Presentor2Y & "," & NextPart.Presentor2Z & "," & NextPart.Presentor2RX & "," & NextPart.Presentor2RY & "," & NextPart.Presentor2RZ & ",0)" 'Location of bit to pick up Screw + 0.1 in the z axis
            Dim removeScrew2Y = NextPart.Presentor2Y - 0.05
            Dim removeScrew2Z = NextPart.Presentor2Z - 0.1
            NextPart.ScrewPres2 = "(" & NextPart.Presentor2X & "," & removeScrew2Y & "," & removeScrew2Z & "," & NextPart.Presentor2RX & "," & NextPart.Presentor2RY & "," & NextPart.Presentor2RZ & ",0)" 'Location for the romoval of the Screw. This type of presenter removes horizontally, so the point is - 0.1 in the y axis
        ElseIf NextPart.screwType = 2 Then
            NextPart.ScrewRemoval = False
        Else
            MsgBox("Unknown screw type. Unavailable")
        End If

        If NextPart.screwType = 2 Then
            screwbot.StartProgram("/programs/FasteningPartsnoscrewv2.urp")
        Else
            screwbot.StartProgram("/programs/FasteningPartsv2.urp")
        End If

        '**************************
        'Lightgate stuff goes here
        '**************************
        Dim bLightGate As Integer = 0
        LightGateCheck("firstchecka", "firstcheckb", NextPart, screwbot, bLightGate)
        If bLightGate = 0 Then
            If NextPart.ScrewRemoval = True Then
                screwbot.WriteTCP("presenta", NextPart.ScrewPresLoc)
                screwbot.WriteTCP("presentb", NextPart.ScrewPres2)
            End If

            bLightGate = 1

            LightGateCheck("screwchecka", "screwcheckb", NextPart, screwbot, bLightGate)
        End If
    End Sub

    Public Sub LightGateCheck(str1 As String, str2 As String, nextpart As NextPart, screwbot As RobotConnect, bLightGate As Integer)
        Dim bScrewCalib = "null"
        If nextpart.screwType = 2 Then
            bScrewCalib = "(1,1)"
        End If

        Try
            While bScrewCalib = "null"
                While screwbot.ReceiveTCP(str1) = False
                End While

                'Commented out for testing without the light gate
                'bLightGate = ReceiveSerialData()
                If bLightGate = 1 Then
                    bScrewCalib = "(1,1)"
                ElseIf bLightGate = 0 Then
                    bScrewCalib = "(0,0)"
                End If
                screwbot.WriteTCP(str2, bScrewCalib)
            End While
        Catch ex As Exception
            MsgBox("Error: ScrewCalibration")
        End Try
    End Sub

    Public Function ReceiveSerialData() As Integer

        Dim Incoming As String

        'Creating the necessary settings when opening com ports - CA
        If bportOpen = True Then
            spComRobot.Close()
            spComRobot.Dispose()
            bportOpen = False
        End If

        If bportOpen = False Then
            spComRobot = New SerialPort("COM4", 9600, Parity.None, 8, StopBits.One)
            spComRobot.Open()
            spComRobot.Handshake = Handshake.None
            spComRobot.Encoding = System.Text.Encoding.Default
            spComRobot.ReadTimeout = 10000
            bportOpen = True
        End If

        'writes to the arduino to read the Screw presence
        spComRobot.Write("done")
Receive:
        Try
            Incoming = spComRobot.ReadLine()

            If Incoming.Length < 1 And Incoming = "done" Then
                sleep(0.5)
                GoTo Receive
            Else
                Dim iIncoming = Int(Incoming)
                If iIncoming = 1 Then
                    Return 1
                ElseIf iIncoming = 0 Then
                    Return 0
                End If
            End If
        Catch ex As TimeoutException
            Return "Error: Serial Port read timed out."
        End Try
        Return Nothing
    End Function
End Module
