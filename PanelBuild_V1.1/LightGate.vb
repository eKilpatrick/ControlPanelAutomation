Imports System.IO.Ports
Module LightGate
    Public Function ReceiveSerialData(ByRef spComRobot As SerialPort, ByRef bportOpen As Boolean) As Boolean

        Dim Incoming As String

        'Creating the necessary settings when opening com ports - CA
        If bportOpen = True Then
            spComRobot.Close()
            spComRobot.Dispose()
            bportOpen = False
        End If

        If bportOpen = False Then
            spComRobot = New SerialPort("COM1", 9600, Parity.None, 8, StopBits.One)
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
                Threading.Thread.Sleep(500)
                GoTo Receive
            Else
                Dim iIncoming = Int(Incoming)
                If iIncoming = 1 Then
                    Return True
                ElseIf iIncoming = 0 Then
                    Return False
                End If
            End If
        Catch ex As TimeoutException
            Return "Error: Serial Port read timed out."
        End Try
        Return Nothing
    End Function
End Module
