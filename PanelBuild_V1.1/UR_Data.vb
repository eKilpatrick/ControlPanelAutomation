Imports System

Namespace RobotIO

    <Serializable>
    Public Class Outputs
        'Public actual_q(5) As Double 'Actual joint positions
        Public actual_TCP_pose(5) As Double 'Actual TCP position
        'Public actual_TCP_speed(5) As Double 'Actual TCP speed
        'Public actual_TCP_force(5) As Double 'Actual TCP force

        Public target_TCP_pose(5) As Double 'Target TCP position
        'Public target_TCP_speed(5) As Double 'Target TCP speed
        'Public target_TCP_force(5) As Double 'Target TCP force
        Public target_speed_fraction As Double

        'Public robot_mode As Int32
        'Public safety_status As Int32
        'Public safety_status_bits As UInt32
        '                                           **************Output Register Uses******************
        '                                           Gripper                      |  Screwbot
        '                                           ----------------------------------------------------
        Public output_bit_register_64 As Boolean '  Program Started              |  Program Started
        Public output_bit_register_65 As Boolean '  Grabbed from Storage         |  At the LightGate
        Public output_bit_register_66 As Boolean '  Placed on Panel              |  At the Presentor
        Public output_bit_register_67 As Boolean '  Back Home                    |  At the LightGate Pt2
        Public output_bit_register_68 As Boolean '  Program Ended                |  Begin Screwing the Part
        Public output_bit_register_69 As Boolean '                               |  Finished Screwing the Part
        Public output_bit_register_70 As Boolean '                               |  Back Home
        Public output_bit_register_71 As Boolean '                               |  Program Ended
        Public output_bit_register_72 As Boolean '                               |  ReAttemptGet

        Public output_double_register_24 As Double ' StorageActual(0)            |  PresActual(0)
        Public output_double_register_25 As Double ' StorageActual(1)            |  PresActual(1)
        Public output_double_register_26 As Double ' StorageActual(2)            |  PresActual(2)
        Public output_double_register_27 As Double ' StorageActual(3)            |  PresActual(3)
        Public output_double_register_28 As Double ' StorageActual(4)            |  PresActual(4)
        Public output_double_register_29 As Double ' StorageActual(5)            |  PresActual(5)
        Public output_double_register_30 As Double ' PanelActual(0)              |  Hole1Actual(0)
        Public output_double_register_31 As Double ' PanelActual(1)              |  Hole1Actual(1)
        Public output_double_register_32 As Double ' PanelActual(2)              |  Hole1Actual(2)
        Public output_double_register_33 As Double ' PanelActual(3)              |  Hole1Actual(3)
        Public output_double_register_34 As Double ' PanelActual(4)              |  Hole1Actual(4)
        Public output_double_register_35 As Double ' PanelActual(5)              |  Hole1Actual(5)
        Public output_double_register_36 As Double '                             |  Hole2Actual(0)
        Public output_double_register_37 As Double '                             |  Hole2Actual(1)
        Public output_double_register_38 As Double '                             |  Hole2Actual(2)
        Public output_double_register_39 As Double '                             |  Hole2Actual(3)
        Public output_double_register_40 As Double '                             |  Hole2Actual(4)
        Public output_double_register_41 As Double '                             |  Hole2Actual(5)
    End Class

    <Serializable>
    Public Class Inputs
        'Public standard_digital_output As Byte
        'Public standard_digital_output_mask As Byte

        Public speed_slider_fraction As Double
        Public speed_slider_mask As UInteger = 1
        '                                           ***********Input Register Uses*****************
        '                                           Gripper                  |  Screwbot
        '                                           --------------------------------------------------
        Public input_double_register_24 As Double = 0 '  StorageBinCoord(0)       |  ScrewPresLoc(0)
        Public input_double_register_25 As Double = 0 '  StorageBinCoord(1)       |  ScrewPresLoc(1)
        Public input_double_register_26 As Double = 0 '  StorageBinCoord(2)       |  ScrewPresLoc(2)
        Public input_double_register_27 As Double = 0 '  StorageBinCoord(3)       |  ScrewPresLoc(3)
        Public input_double_register_28 As Double = 0 '  StorageBinCoord(4)       |  ScrewPresLoc(4)
        Public input_double_register_29 As Double = 0 '  StorageBinCoord(5)       |  ScrewPresLoc(5)
        Public input_double_register_30 As Double = 0 '  PosX                     |  Screw1PosAbovePanel(0)
        Public input_double_register_31 As Double = 0 '  PosY                     |  Screw1PosAbovePanel(1)
        Public input_double_register_32 As Double = 0 '  ZPlace                   |  Screw1PosAbovePanel(2)
        Public input_double_register_33 As Double = 0 '  arrPanelOrigin(3)        |  Screw1PosAbovePanel(3)
        Public input_double_register_34 As Double = 0 '  arrPanelOrigin(4)        |  Screw1PosAbovePanel(4)
        Public input_double_register_35 As Double = 0 '  arrPanelOrigin(5)        |  Screw1PosAbovePanel(5)
        Public input_double_register_36 As Double = 0 '  Mass                     |  Screw2PosAbovePanel(0)
        Public input_double_register_37 As Double = 0 '                           |  Screw2PosAbovePanel(1)
        Public input_double_register_38 As Double = 0 '                           |  Screw2PosAbovePanel(2)
        Public input_double_register_39 As Double = 0 '                           |  Screw2PosAbovePanel(3)
        Public input_double_register_40 As Double = 0 '                           |  Screw2PosAbovePanel(4)
        Public input_double_register_41 As Double = 0 '                           |  Screw2PosAbovePanel(5)
        Public input_double_register_42 As Double = 0 '                           |  Depth
        Public input_double_register_43 As Double = 0 '                           |  Offset
        'Public input_double_register_44 As Double = 0 '
        'Public input_double_register_45 As Double = 0 '
        'Public input_double_register_46 As Double = 0 '
        'Public input_double_register_47 As Double = 0 '

        Public input_bit_register_64 As Boolean = False '   WaitToLeave              |  ScrewRemoval
        Public input_bit_register_65 As Boolean = False '                            |  tbOffset
        Public input_bit_register_66 As Boolean = False '                            |  WaitToScrew
        Public input_bit_register_67 As Boolean = False '                            |  WhichScrew - As in the first or second screw for a part, not which type of screw
        Public input_bit_register_68 As Boolean = False '                            |  WhichPres
        Public input_bit_register_69 As Boolean = False '                            |  LightGateBool
        Public input_bit_register_70 As Boolean = False '                            |  LightGate1
        Public input_bit_register_71 As Boolean = False '                            |  LightGate2

    End Class
End Namespace
