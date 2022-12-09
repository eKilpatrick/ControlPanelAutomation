Public Class Part
    Public drawingNum As String
    Public sPanel As String

    Public sPN As String
    Public splitPN As String

    'Part Location from Y_Panel_Build_Positions
    Public X1 As Double
    Public Y1 As Double
    Public Z1 As Double
    Public X2 As Double
    Public Y2 As Double
    Public Z2 As Double
    Public SizeX As Double
    Public SizeY As Double
    Public SizeZ As Double
    Public CenterX As Double
    Public CenterY As Double
    Public PartRotate As Boolean

    'Storage Location from Y_Panel_Build_Storage_Location
    Public BinNum As Integer
    Public StorageX As Double
    Public StorageY As Double
    Public BinStatus As Integer
    Public CommonName As String
    Public StorageZ As Double
    Public StorageRX As Double
    Public StorageRY As Double
    Public StorageRZ As Double
    Public StorageBinCoord As Double() 'Coordinates to grab part in storage bin

    'Panel Origins based on cart from Y_Panel_Build_Origins
    Public panelX As Double
    Public panelY As Double
    Public panelZ As Double
    Public panelRX As Double
    Public panelRY As Double
    Public panelRZ As Double
    Public arrPanelOrigin As Double()

    'Sbot Origins based on cart from Y_Panel_Build_Origins
    Public sbotX As Double
    Public sbotY As Double
    Public sbotZ As Double
    Public sbotRX As Double
    Public sbotRY As Double
    Public sbotRZ As Double
    Public arrSBot As Double()

    'Part size from Y_Panel_Build_Partsize
    Public PartSizeX As Double
    Public PartSizeY As Double
    Public PartSizeZ As Double
    Public HX1 As Double
    Public HY1 As Double
    Public HX2 As Double
    Public HY2 As Double
    Public Mass As Double
    Public ScrewType As String

    'Added for OrientPlacePart
    Public HoleDist1 As Double
    Public HoleDist2 As Double
    Public PHoleX As Double
    Public PHoleY As Double
    Public ZPanel As Double
    Public ZPlace As Double
    Public PosX As Double
    Public PosY As Double
    Public PosZ As Double
    Public AbovePanel As Double()
    Public PlacePanel As Double()
    Public WaitToLeave As Boolean

    'Added for Screw Utilities
    Public HoleX1 As Double
    Public HoleY1 As Double
    Public HoleX2 As Double
    Public HoleY2 As Double
    Public ScrewPresLoc As Double()
    Public ScrewRemoval As Boolean
    Public bLightGate As Integer
    Public HoleRX As Double
    Public HoleRY As Double
    Public Screw1PosAbovePanel As Double()
    Public Screw2PosAbovePanel As Double()
    Public Offset As Double
    Public Depth As Double
    Public WaitToScrew As Boolean = True
    Public WhichScrew As Boolean
    Public WhichPres As Boolean
    Public tbOffset As Boolean = False

    Public Sub InitializeVariables(sPN As String, splitPN As String, drawingNum As String, screwbot As RobotConnect)
#Region "Part Data Initialization"
        Me.drawingNum = drawingNum
        Me.sPN = sPN
        Me.splitPN = splitPN

        'Y_Panel_Build_Positions Panel PartNumber
        sPanel = GetPanel(Me.drawingNum)

        'Y_Panel_Build_Positions data initialization
        Dim PartLoc As DataRow = GetLocation(drawingNum, sPN)
        X1 = CDbl(PartLoc.Item(2)) * 0.0254
        Y1 = CDbl(PartLoc.Item(3)) * 0.0254
        Z1 = CDbl(PartLoc.Item(4)) * 0.0254
        X2 = CDbl(PartLoc.Item(5)) * 0.0254
        Y2 = CDbl(PartLoc.Item(6)) * 0.0254
        Z2 = CDbl(PartLoc.Item(8)) * 0.0254
        SizeX = Math.Abs(X2 - X1)
        SizeY = Math.Abs(Y2 - Y1)
        SizeZ = Math.Abs(Z2 - Z1)
        CenterX = (X1 + X2) / 2
        CenterY = (Y1 + Y2) / 2

        'Y_Panel_Build_Storage_Location data initialization
        Dim StorLoc As DataRow = GetBinCoordinates(splitPN)
        BinNum = StorLoc.Item(1)
        StorageX = StorLoc.Item(2)
        StorageY = StorLoc.Item(3)
        BinStatus = StorLoc.Item(4)
        CommonName = StorLoc.Item(6)
        StorageZ = StorLoc.Item(7)
        StorageRX = StorLoc.Item(8)
        StorageRY = StorLoc.Item(9)
        StorageRZ = StorLoc.Item(10)

        'Y_Panel_Build_Origins Panel data initialization
        Dim PanelOrigin As DataRow = GetOrigin("CART1")
        panelX = PanelOrigin.Item(1)
        panelY = PanelOrigin.Item(2)
        panelZ = PanelOrigin.Item(3)
        panelRX = PanelOrigin.Item(4)
        panelRY = PanelOrigin.Item(5)
        panelRZ = PanelOrigin.Item(6)
        arrPanelOrigin = {panelX, panelY, panelZ, panelRX, panelRY, panelRZ}

        'Y_Panel_Build_Origins Sbot data initialization
        Dim SbotOrigin As DataRow = GetOrigin("SCREWBOT OFFSET")
        sbotX = panelX + SbotOrigin.Item(1)
        sbotY = panelY + SbotOrigin.Item(2)
        sbotZ = panelZ + SbotOrigin.Item(3)
        sbotRX = panelRX + SbotOrigin.Item(4)
        sbotRY = panelRY + SbotOrigin.Item(5)
        sbotRZ = panelRZ + SbotOrigin.Item(6)
        arrSBot = {sbotX, sbotY, sbotZ, sbotRX, sbotRY, sbotRZ}

        'Y_Panel_Build_Partsize data initialization
        Dim PartSize As DataRow = GetPartSize(splitPN)
        PartSizeX = CDbl(PartSize.Item(1) * 0.0254)
        PartSizeY = CDbl(PartSize.Item(2) * 0.0254)
        PartSizeZ = CDbl(PartSize.Item(3) * 0.0254)
        HX1 = CDbl(PartSize.Item(4) * 0.0254)
        HY1 = CDbl(PartSize.Item(5) * 0.0254)
        HX2 = CDbl(PartSize.Item(6) * 0.0254)
        HY2 = CDbl(PartSize.Item(7) * 0.0254)
        ScrewType = PartSize.Item(8)
        Mass = PartSize.Item(10)
#End Region

#Region "Gripper Utilities"
        'Changes where to grab the part based on its size
        GripperOrientation(Me)

        'Storage bin coordinates array initialization
        StorageBinCoord = {StorageX, StorageY, StorageZ, StorageRX, StorageRY, StorageRZ}

        'Changes the origin and rotation of the part to be placed on the panel
        OrientPlacePart(sPanel, Me)

        'If it's one of these parts, the gripper should place it and leave immediately
        If CommonName = "CIRCUIT BREAKER" Or CommonName = "KNIFESWITCH" Or CommonName = "FUSEBLOCK" Then
            WaitToLeave = False
        Else
            WaitToLeave = True
        End If

        'Locations of where the part is to be placed above/on the panel in the panel plane.
        PlacePanel = {PosX, PosY, ZPlace, arrPanelOrigin(3), arrPanelOrigin(4), arrPanelOrigin(5)}
#End Region
#Region "Screwbot Utilities"
        'Sets the part to active in Y_Panel_Build_Robot_Offsets
        SendActivePart(splitPN)

        'Sets the origin to the ACTUAL origin of the panel based on the distance from the top mounting hole using the dimensions in the database
        SetMountOrigin(Me)

        'Calculates the Hole locations on the panel based on part rotation and part size
        ScrewPositionCalibration(Me)

        'Initializes the Screw Presentor information based on the screw type
        SetActiveScrewPresentor(Me, screwbot)

        'If the part is a terminal block, the screwbot starts outside of part, goes down, and then slides into the hole...
        TBScrewFix(Me, screwbot)

        If PartRotate = True Then
            HoleRX = 2.2
            HoleRY = 2.2
        Else
            HoleRX = 0
            HoleRY = -3.14
        End If

        'Location of screwing above the panel
        Screw1PosAbovePanel = {HoleX1, HoleY1, 0.2, HoleRX, HoleRY, 0.0}
        Screw2PosAbovePanel = {HoleX2, HoleY2, 0.2, HoleRX, HoleRY, 0.0}

        Dim ActiveOffsetDepth As DataRow = GetDepthOffset()
        Depth = ActiveOffsetDepth.Item(0)
        Offset = ActiveOffsetDepth.Item(1)

#End Region
    End Sub

    Public Sub SendVarsRobot(ByRef screwbot As RobotConnect, ByRef gripper As RobotConnect)
        gripper.UrInputs.input_double_register_24 = StorageBinCoord(0)
        gripper.UrInputs.input_double_register_25 = StorageBinCoord(1)
        gripper.UrInputs.input_double_register_26 = StorageBinCoord(2)
        gripper.UrInputs.input_double_register_27 = StorageBinCoord(3)
        gripper.UrInputs.input_double_register_28 = StorageBinCoord(4)
        gripper.UrInputs.input_double_register_29 = StorageBinCoord(5)

        gripper.UrInputs.input_double_register_30 = PlacePanel(0)
        gripper.UrInputs.input_double_register_31 = PlacePanel(1)
        gripper.UrInputs.input_double_register_32 = PlacePanel(2)
        gripper.UrInputs.input_double_register_33 = PlacePanel(3)
        gripper.UrInputs.input_double_register_34 = PlacePanel(4)
        gripper.UrInputs.input_double_register_35 = PlacePanel(5)

        gripper.UrInputs.input_double_register_36 = Mass

        gripper.UrInputs.input_bit_register_64 = WaitToLeave

        gripper.UrInputs.speed_slider_fraction = 0.75

        Dim gripperUpdated As Boolean = gripper.RTDE_Client.Send_Ur_Inputs()
        If gripperUpdated Then
            Console.WriteLine("Gripper registers successfully updated")
        Else
            Console.WriteLine("Gripper registers NOT updated")
        End If

        screwbot.UrInputs.input_double_register_24 = ScrewPresLoc(0)
        screwbot.UrInputs.input_double_register_25 = ScrewPresLoc(1)
        screwbot.UrInputs.input_double_register_26 = ScrewPresLoc(2)
        screwbot.UrInputs.input_double_register_27 = ScrewPresLoc(3)
        screwbot.UrInputs.input_double_register_28 = ScrewPresLoc(4)
        screwbot.UrInputs.input_double_register_29 = ScrewPresLoc(5)

        screwbot.UrInputs.input_double_register_30 = Screw1PosAbovePanel(0)
        screwbot.UrInputs.input_double_register_31 = Screw1PosAbovePanel(1)
        screwbot.UrInputs.input_double_register_32 = Screw1PosAbovePanel(2)
        screwbot.UrInputs.input_double_register_33 = Screw1PosAbovePanel(3)
        screwbot.UrInputs.input_double_register_34 = Screw1PosAbovePanel(4)
        screwbot.UrInputs.input_double_register_35 = Screw1PosAbovePanel(5)

        screwbot.UrInputs.input_double_register_36 = Screw2PosAbovePanel(0)
        screwbot.UrInputs.input_double_register_37 = Screw2PosAbovePanel(1)
        screwbot.UrInputs.input_double_register_38 = Screw2PosAbovePanel(2)
        screwbot.UrInputs.input_double_register_39 = Screw2PosAbovePanel(3)
        screwbot.UrInputs.input_double_register_40 = Screw2PosAbovePanel(4)
        screwbot.UrInputs.input_double_register_41 = Screw2PosAbovePanel(5)

        screwbot.UrInputs.input_double_register_42 = Depth
        screwbot.UrInputs.input_double_register_43 = Offset

        screwbot.UrInputs.input_bit_register_64 = ScrewRemoval
        screwbot.UrInputs.input_bit_register_65 = tbOffset
        screwbot.UrInputs.input_bit_register_66 = WaitToScrew
        'screwbot.urinputs.input_bit_register_67 = WhichScrew is set whenever you call the ScrewPart() function so it can be a dynamic choice rather than set.
        screwbot.UrInputs.input_bit_register_68 = WhichPres

        screwbot.UrInputs.speed_slider_fraction = 0.75

        Dim screwbotUpdated As Boolean = screwbot.RTDE_Client.Send_Ur_Inputs()
        If screwbotUpdated Then
            Console.WriteLine("Screwbot registers successfully updated")
        Else
            Console.WriteLine("Screwbot registers NOT updated")
        End If
    End Sub
End Class
