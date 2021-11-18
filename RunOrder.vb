Module RunOrder
    Public Sub RunOrder(DN As String, ByRef screwbot As RobotConnect, ByRef gripbot As RobotConnect)
        Dim KnifeSwitches As NextPart() = {}
        Try
            screwbot.Connect()
            gripbot.Connect()
        Catch ex As Exception
            MsgBox("There was an error trying to connect to the robots. If your computer isn't connected to them, simulate the robots")
            Exit Sub
        End Try

        While True

            Dim NextPart As New NextPart(DN)
            Dim PartBool As Boolean = NextPart.InitializePartVariables()
            If PartBool = False Then
                Exit While
            End If

            'Storage bin stuff
            If NextPart.commonName = "FUSEBLOCK" Or NextPart.commonName = "CIRCUIT BREAKER" Or NextPart.commonName = "KNIFESWITCH" Then
                gripbot.StartProgram("/programs/PickTallPart.urp")
            Else
                gripbot.StartProgram("/programs/PickPart.urp")
            End If
            'If NextPart.PartRotate = True Then
            'NextPart.storageApproach = "(" & NextPart.storageX & "," & NextPart.storageY & "," & NextPart.storageZ + 0.057 & "," & NextPart.storageRX & "," & NextPart.storageRY & "," & NextPart.storageRZ & ",0)"
            'NextPart.storagePos = "(" & NextPart.storageX & "," & NextPart.storageY & "," & NextPart.storageZ & "," & NextPart.storageRX & "," & NextPart.storageRY & "," & NextPart.storageRZ & ",0)"
            'Else
            NextPart.storageApproach = "(" & NextPart.storageX + 0.077 & "," & NextPart.storageY & "," & NextPart.storageZ + 0.057 & "," & NextPart.storageRX & "," & NextPart.storageRY & "," & NextPart.storageRZ & ",0)"
            NextPart.storagePos = "(" & NextPart.storageX & "," & NextPart.storageY & "," & NextPart.storageZ & "," & NextPart.storageRX & "," & NextPart.storageRY & "," & NextPart.storageRZ & ",0)"
            'End If

            gripbot.WriteTCP("ApproachBin", NextPart.storageApproach)
            gripbot.WriteTCP("BinCoords", NextPart.storagePos)

            'DatabaseQueries.ChangeStorage(NextPart.storageSlot, -1)

            'Place Part stuff
            If NextPart.PartRotate = True Then
                NextPart.panelApproach = "(" & (NextPart.X1 + NextPart.sizeY / 2) & "," & (NextPart.Y1 + NextPart.sizeX / 2) & "," & (NextPart.sizeZ + 0.1) & ",-2.22,2.22,0.0,0)"
                NextPart.panelPos = "(" & (NextPart.X1 + NextPart.sizeY / 2) & "," & (NextPart.Y1 + NextPart.sizeX / 2) & "," & (NextPart.sizeZ - 0.035) & ",-2.22,2.22,0.0,0)"

            Else
                NextPart.panelApproach = "(" & (NextPart.X1 + NextPart.sizeX / 2) & "," & (NextPart.Y1 + NextPart.sizeY / 2) & "," & (NextPart.sizeZ + 0.1) & ",3.14,0.0,0.0,0)"
                NextPart.panelPos = "(" & (NextPart.X1 + NextPart.sizeX / 2) & "," & (NextPart.Y1 + NextPart.sizeY / 2) & "," & (NextPart.sizeZ - 0.035) & ",3.14,0.0,0.0,0)"
            End If

            gripbot.WriteTCP("GoPanel", NextPart.panelApproach)
            gripbot.WriteTCP("PlacePanel", NextPart.panelPos)

            Dim currentProg As String = gripbot.GetLoadedProgram()
            If currentProg.Contains("PickPart.urp") Then
                gripbot.Load("/programs/PickPartv2part2.urp")
            End If

            ScrewProcess.StartScrewing(NextPart, screwbot, NextPart.ksPt2)

            DatabaseQueries.SendComplete(NextPart.PN, NextPart.DN)

            If NextPart.commonName = "KNIFESWITCH" Then
                KnifeSwitches.Append(NextPart)
            End If

        End While

        If KnifeSwitches.Count <> 0 Then
            Dim knifeOpen = False
            While knifeOpen = False
                knifeOpen = MsgBox("Please lift up all the knifeswitches in order to put the second screw", MessageBoxButtons.YesNo, "KNIFESWITCH ACTION REQUIRED")
            End While
            For counter As Integer = 0 To (KnifeSwitches.Count - 1)
                KnifeSwitches(counter).ksPt2 = True
                ScrewProcess.StartScrewing(KnifeSwitches(counter), screwbot, KnifeSwitches(counter).ksPt2)
            Next
        End If

        MsgBox("No parts remaining")
    End Sub
End Module
