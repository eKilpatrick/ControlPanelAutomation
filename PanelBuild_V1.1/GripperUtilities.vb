Module GripperUtilities
    Public GripperLength As Double = 0.066675
    Public ZGripper As Double = 0.0619125
    Public Theta As Double = (34 * Math.PI / 180)
    'Affects the storage location of the part
    Public Sub GripperOrientation(ByRef Part As Part)
        Dim DeltaX As Double
        Dim DeltaZ As Double
        If Part.CommonName = "12PT TB" Or Part.CommonName = "POWER TB" Then
            DeltaX = (Part.PartSizeY / 2) * (Math.Sin(Theta))
            DeltaZ = (Part.PartSizeY / 2) * (Math.Cos(Theta))
        End If

        If Part.CommonName = "CIRCUIT BREAKER" Or Part.CommonName = "KNIFESWITCH" Or Part.CommonName = "FUSEBLOCK" Then
            Part.StorageX = Part.StorageX
            Part.StorageZ = Part.StorageZ
        Else
            'If you don't want to grab the part from the predetermined location in the database
            'Currently only used for Terminal Blocks and Power Blocks
            Part.StorageX = Part.StorageX - DeltaX
            Part.StorageZ = Part.StorageZ + DeltaZ
        End If
    End Sub

    'Affects the panel location of the part
    Public Sub OrientPlacePart(sPanel As String, ByRef Part As Part)
        Part.PartRotate = False

        Part.HoleDist1 = Math.Sqrt((((Part.SizeX / 2) - Part.HX1) ^ 2) + (((Part.SizeY / 2) - Part.HY1) ^ 2))
        Part.HoleDist2 = Math.Sqrt((((Part.SizeX / 2) - Part.HX2) ^ 2) + (((Part.SizeY / 2) - Part.HY2) ^ 2))

        If Part.SizeY < 0.1522 And (Part.HoleDist1 < 0.0889) Or Part.SizeY < 0.1522 And (Part.HoleDist2 < 0.0889) Then
            Part.SizeY += (GripperLength * 2)
        End If

        Dim PanelHole As DataRow = GetPanelHole(sPanel)
        Part.PHoleX = CDbl(PanelHole.Item(4)) * 0.0254
        Part.PHoleY = CDbl(PanelHole.Item(5)) * 0.0254

        If Part.PHoleX < 0.05555 Then
            Part.arrPanelOrigin(0) -= Part.PHoleX
        Else
            Part.arrPanelOrigin(0) -= 0.04675
        End If

        Part.arrPanelOrigin(1) -= Part.PHoleY

        Part.ZPanel = Part.arrPanelOrigin(2) + 0.003                     'NEW CODE HERE!!! HARCODED Z DISPLACEMENT VALUES BY PART FOR GRIPPER
        If Part.CommonName = "FUSEBLOCK" Then
            Part.ZPlace = Part.ZPanel - 0.01439 ' - 0.02139
        ElseIf Part.CommonName = "KNIFESWITCH" Then
            Part.ZPlace = Part.ZPanel - 84 '- 0.00784
        ElseIf Part.CommonName = "CIRCUIT BREAKER" Then
            '//TODO:THIS NEEDS TO BE RECALCULATED
            Part.ZPlace = Part.ZPanel + Part.PartSizeZ - ZGripper - 0.00331 '- 0.00831
        ElseIf Part.CommonName = "12PT TB" Then
            Part.ZPlace = Part.ZPanel - 0.009 '- 0.00464
        Else
            Part.ZPlace = Part.ZPanel
        End If

        If Part.CommonName = "KNIFESWITCH" Then               'NEW CODE HERE!!! HARDCODED X/Y OFFSETS BY PART FOR GRIPPER
            Part.PosX = CDbl(Part.arrPanelOrigin(0)) + Part.CenterX - 0.0016 + 0.00804
            Part.PosY = CDbl(Part.arrPanelOrigin(1)) + Part.CenterY + 0.0005 - 0.00326
            Part.PosZ = Part.ZPlace + 0.2
        ElseIf Part.CommonName = "CIRCUIT BREAKER" Then
            Part.PosX = CDbl(Part.arrPanelOrigin(0)) + Part.CenterX + 0.0121 + 0.03754
            Part.PosY = CDbl(Part.arrPanelOrigin(1)) + Part.CenterY + 0.001 + 0.00311
            Part.PosZ = Part.ZPlace + 0.2
        ElseIf Part.CommonName = "FUSEBLOCK" Then
            Part.PosX = CDbl(Part.arrPanelOrigin(0)) + Part.CenterX - 0.0027 + 0.02115
            Part.PosY = CDbl(Part.arrPanelOrigin(1)) + Part.CenterY + 0.0006 + 0.00504
            Part.PosZ = Part.ZPlace + 0.2
        ElseIf Part.CommonName = "12PT TB" Then
            Part.PosX = CDbl(Part.arrPanelOrigin(0)) + Part.CenterX - 0.0018 - 0.00253
            Part.PosY = CDbl(Part.arrPanelOrigin(1)) + Part.CenterY + 0.0005 + 0.0025
            Part.PosZ = Part.ZPlace + 0.2
            Part.arrPanelOrigin(3) += 0.006
            Part.arrPanelOrigin(4) -= 0.011      'Added to make holes match on either side of the terminal block (needed to be rotated for some reason)
        Else
            Part.PosX = CDbl(Part.arrPanelOrigin(0)) + Part.CenterX
            Part.PosY = CDbl(Part.arrPanelOrigin(1)) + Part.CenterY
            Part.PosZ = Part.ZPlace + 0.2
        End If

        If Part.CommonName = "FUSEBLOCK" Or Part.CommonName = "KNIFESWITCH" Or Part.CommonName = "CIRCUIT BREAKER" Then
            Part.PartRotate = True
            Part.arrPanelOrigin(3) = -2.22         'rotates the part CCW 90 degrees
            Part.arrPanelOrigin(4) = 2.22
            If Part.CommonName = "KNIFESWITCH" Then
                Part.PosX -= GripperLength / 2
                'TODO:// this is hacky
            ElseIf Part.CommonName = "FUSEBLOCK" Then
                Part.PosX -= GripperLength * 0.43
            Else
                '//ERROR COULD OCCUR HERE WITH CIRCUIT BREAKER
                Part.PosX -= GripperLength
            End If
        Else
            '//do nothing yet.. something could be added depending on the parts
        End If
    End Sub

    'This function resets the status bits on the robots
    Public Sub ResetOutputVars(ByRef robot As RobotConnect)
        robot.StartProgram("/programs/ResetVars.urp")
    End Sub

    'Starts the Gripper program NewPickPart.urp
    Public Sub GetPart(ByRef gripper As RobotConnect)
        gripper.StartProgram("/programs/_NewPickPart.urp")
        WaitProgramStart(gripper)
    End Sub

    Public Sub LeavePanel(ByRef gripper As RobotConnect)
        gripper.UrInputs.input_bit_register_64 = False
        If gripper.RTDE_Client.Send_Ur_Inputs Then
            Console.WriteLine("Successfully udpated gripper's registers")
        Else
            Console.WriteLine("Error updating gripper's registers")
        End If
    End Sub

#Region "Gripper Status Updates"
    Public Sub WaitProgramStart(ByRef gripper As RobotConnect)
        While gripper.UrOutputs.output_bit_register_64 <> True
            'Wait until program has started
            Threading.Thread.Sleep(100)
        End While
    End Sub

    Public Sub GrabbedFromStorage(ByRef gripper As RobotConnect)
        While gripper.UrOutputs.output_bit_register_65 <> True
            'Wait until the robot has grabbed the part from storage
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub PlacedOnPanel(ByRef gripper As RobotConnect)
        While gripper.UrOutputs.output_bit_register_66 <> True
            'Wait until the robot has placed the part on the panel
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub GoingBackHome(ByRef gripper As RobotConnect)
        While gripper.UrOutputs.output_bit_register_67 <> True
            'Wait until the robot is going back home
            Threading.Thread.Sleep(250)
        End While
    End Sub

    Public Sub ProgramEnded(ByRef gripper As RobotConnect)
        While gripper.UrOutputs.output_bit_register_68 <> True
            'Wait until the robot program has completed
            Threading.Thread.Sleep(250)
        End While
    End Sub
#End Region
End Module
