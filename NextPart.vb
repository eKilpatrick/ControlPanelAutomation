Public Class NextPart
    Public DN As String
    Public ColorStatus As Color = Color.Red

    Public PN As String
    Public X1 As Double
    Public Y1 As Double
    Public Z1 As Double
    Public X2 As Double
    Public Y2 As Double
    Public Z2 As Double
    Public splitPN As String

    Public storageSlot As Integer
    Public storageX As Double
    Public storageY As Double
    Public commonName As String
    Public storageZ As Double
    Public storageRX As Double
    Public storageRY As Double
    Public storageRZ As Double

    Public sizeX As Double
    Public sizeY As Double
    Public sizeZ As Double
    Public hole1X As Double
    Public hole1Y As Double
    Public hole2X As Double
    Public hole2Y As Double
    Public screwType


    Public PanelX As Double
    Public PanelY As Double
    Public PanelZ As Double
    Public PanelRX As Double
    Public PanelRY As Double
    Public PanelRZ As Double


    Public sbotX As Double
    Public sbotY As Double
    Public sbotZ As Double
    Public sbotRX As Double
    Public sbotRY As Double
    Public sbotRZ As Double

    Public Presentor1X As Double
    Public Presentor1Y As Double
    Public Presentor1Z As Double
    Public Presentor1RX As Double
    Public Presentor1RY As Double
    Public Presentor1RZ As Double

    Public Presentor2X As Double
    Public Presentor2Y As Double
    Public Presentor2Z As Double
    Public Presentor2RX As Double
    Public Presentor2RY As Double
    Public Presentor2RZ As Double

    Public Depth As Double
    Public Offset As Double

    Public PartRotate As Boolean
    'Some of these variables aren't needed with the new calculation method

    Public Screw1 As Boolean = False
    Public Screw2 As Boolean = False
    'Public HoleX1 As Double
    'Public HoleX2 As Double
    'Public HoleY1 As Double
    'Public HoleY2 As Double
    Public ScrewPresLoc As String
    Public ScrewPres2 As String
    Public ScrewRemoval As Boolean = True
    Public ScrewPos1 As String
    Public ScrewPos2 As String
    Public ScrewPos3 As String
    Public tbOffset As Double
    Public ksPt2 As Boolean = False

    'New variables
    Public storagePos As String
    Public storageApproach As String
    Public panelPos As String
    Public panelApproach As String

    Public Sub New(DN As String)
        Me.DN = DN
    End Sub

    Public Function InitializePartVariables()
        Dim partRow As DataRow = DatabaseQueries.GetFirstItem(DN)

        If partRow Is Nothing Then
            Return False
        Else
            'Gets the part number and panel location of the top item, converting location to meters
            PN = partRow(1)
            X1 = (CDbl(partRow(2)) - 11) * 0.0254
            Y1 = (CDbl(partRow(3)) - 1.5) * 0.0254
            Z1 = CDbl(partRow(4)) * 0.0254
            X2 = (CDbl(partRow(5)) - 11) * 0.0254
            Y2 = (CDbl(partRow(6)) - 1.5) * 0.0254
            Z2 = CDbl(partRow(7)) * 0.0254

            If PN.Contains("_") Then
                splitPN = PN.Split("_")(0)
            Else
                splitPN = PN.Split(":")(0)
            End If

            'Gets the storage location where that part is stored
            Dim storageRow As DataRow = DatabaseQueries.GetStorageLocation(splitPN)
            If storageRow Is Nothing Then
                MsgBox(splitPN & " could not be found in the storage bin. Please verify it is there or add more.")
                Return False
            Else
                storageSlot = storageRow(0)
                storageX = storageRow(3)
                storageY = storageRow(4)
                commonName = storageRow(9)
                storageZ = storageRow(5)
                storageRX = storageRow(6)
                storageRY = storageRow(7)
                storageRZ = storageRow(8)
            End If

            'Gets the size, hole locations, and screw type of the part, converting to meters
            Dim partSizeRow As DataRow = DatabaseQueries.GetPartSize(splitPN)
            If partSizeRow Is Nothing Then
                MsgBox(splitPN & " doesn't have a partsize or is an incorrect part")
                Return False
            Else
                sizeX = CDbl(partSizeRow(1)) * 0.0254
                sizeY = CDbl(partSizeRow(2)) * 0.0254
                sizeZ = CDbl(partSizeRow(3)) * 0.0254
                hole1X = CDbl(partSizeRow(4)) * 0.0254
                hole1Y = CDbl(partSizeRow(5)) * 0.0254
                hole2X = CDbl(partSizeRow(6)) * 0.0254
                hole2Y = CDbl(partSizeRow(7)) * 0.0254
                screwType = partSizeRow(8)
            End If

            'Gets the origins of the cart and the screwbot
            'Currently using cart 1 as the default every time...
            Dim panelOriginRow As DataRow = DatabaseQueries.GetOrigin("CART1")
            If panelOriginRow Is Nothing Then
                MsgBox("There is no row of information for the top left mounting screw in the database table Origins")
                Return False
            Else
                PanelX = panelOriginRow(1)
                PanelY = panelOriginRow(2)
                PanelZ = panelOriginRow(3)
                PanelRX = panelOriginRow(4)
                PanelRY = panelOriginRow(5)
                PanelRZ = panelOriginRow(6)
            End If

            Dim sbotOriginRow As DataRow = DatabaseQueries.GetOrigin("SCREWBOT OFFSET")
            If sbotOriginRow Is Nothing Then
                MsgBox("There is no row of information for the screwbot offsets in the database table Origins")
                Return False
            Else
                sbotX = sbotOriginRow(1) + panelOriginRow(1)
                sbotY = sbotOriginRow(2) + panelOriginRow(2)
                sbotZ = sbotOriginRow(3) + panelOriginRow(3)
                sbotRX = sbotOriginRow(4)
                sbotRY = sbotOriginRow(5)
                sbotRZ = sbotOriginRow(6)
            End If

            Dim Presentor1Origin As DataRow = DatabaseQueries.GetOrigin("PRESENTOR1")
            If Presentor1Origin Is Nothing Then
                MsgBox("There is no information for presentor 1's origin in the database table Origins")
                Return False
            Else
                Presentor1X = Presentor1Origin.Item(1)
                Presentor1Y = Presentor1Origin.Item(2)
                Presentor1Z = Presentor1Origin.Item(3)
                Presentor1RX = Presentor1Origin.Item(4)
                Presentor1RY = Presentor1Origin.Item(5)
                Presentor1RZ = Presentor1Origin.Item(6)
            End If

            Dim Presentor2Origin As DataRow = DatabaseQueries.GetOrigin("PRESENTOR2")
            If Presentor2Origin Is Nothing Then
                MsgBox("There is no information for presentor 2's origin in the database table Origins")
                Return False
            Else
                Presentor2X = Presentor2Origin.Item(1)
                Presentor2Y = Presentor2Origin.Item(2)
                Presentor2Z = Presentor2Origin.Item(3)
                Presentor2RX = Presentor2Origin.Item(4)
                Presentor2RY = Presentor2Origin.Item(5)
                Presentor2RZ = Presentor2Origin.Item(6)
            End If

            'Gets the offset and depth for the specific part
            Dim RobotOffsets As DataRow = DatabaseQueries.GetRobotOffsets(splitPN)
            If RobotOffsets Is Nothing Then
                MsgBox("There is no information regarding the depth and offset for the part: " & PN)
                Return False
            Else
                Depth = RobotOffsets.Item(2)
                Offset = RobotOffsets.Item(4)
            End If

            If (Math.Round(10000 * sizeX) / 10000) <> (Math.Round((10000 * (X2 - X1))) / 10000) Then
                PartRotate = True
            Else
                PartRotate = False
            End If

            Return True
        End If
    End Function

End Class



