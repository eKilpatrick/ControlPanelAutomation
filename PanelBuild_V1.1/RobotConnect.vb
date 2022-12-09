Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Public Class RobotConnect

    'Public attributes that each object of the class will use
    Public name As String
    Public connected As Boolean = False
    Public simulate As Boolean = False

    Public ipAddress As IPAddress
    Public serverIPaddress As IPAddress
    Public robotSock As Socket
    Public listeningPort As Integer

    Public RTDE_Client As New Ur_Rtde.RtdeClient()
    Public UrOutputs As New RobotIO.Outputs()
    Public UrInputs As New RobotIO.Inputs()

    'Constructor that sets the ip address of each object
    Public Sub New(name As String, ip As String, ip2 As String, port As Integer)
        Me.name = name
        ipAddress = IPAddress.Parse(ip)
        serverIPaddress = IPAddress.Parse(ip2)
        listeningPort = port
    End Sub

    'Destructor that closes socket connections upon end of the program
    Protected Overrides Sub Finalize()
        If connected = True Then
            robotSock.Close()
            RTDE_Client.Disconnect(True)
        End If
    End Sub

    'Creates a socket connection to the specified robot dashboard port
    Public Sub Connect()
        If simulate = True Then
            Exit Sub
        End If
        Dim robotEP As New IPEndPoint(ipAddress, 29999)
        Dim s As New Socket(ipAddress.AddressFamily, SocketType.Stream, ProtocolType.Tcp)
        Dim ConnectError As Integer = 0
        While connected = False
            If ConnectError = 5 Then
                MsgBox("The connection to the robots has errored out 5 times, please try again later")
                Exit Sub
            End If
            Try
                s.Connect(robotEP)
                connected = True
                robotSock = s
            Catch ex As Exception
                ConnectError += 1
                Threading.Thread.Sleep(500)
            End Try
        End While

        Try
            Dim bytes(1024) As Byte
            Dim bytesRec As Integer = s.Receive(bytes)
            If bytesRec <> 0 Then
                Dim strRec = Encoding.ASCII.GetString(bytes, 0, bytesRec).Trim()
                Console.WriteLine(strRec)
            End If
        Catch ex As Exception
            MsgBox("Error receiving information after connecting to the robot: " & ex.Message)
        End Try

        Try
            AddHandler RTDE_Client.OnSockClosed, AddressOf Robot_OnSockClosed
            If RTDE_Client.Connect(ipAddress.ToString, 2) Then
                Console.WriteLine("Connected to rtde: " & name)
            Else
                Console.WriteLine("Failed to connect to rtde: " & name)
            End If

            Dim InputSuccess = RTDE_Client.Setup_Ur_Inputs(UrInputs)
            Dim OutputSuccess = RTDE_Client.Setup_Ur_Outputs(UrOutputs, 10)

            AddHandler RTDE_Client.OnDataReceive, AddressOf Robot_OnDataReceive
            Dim rtde_connected As Boolean = RTDE_Client.Ur_ControlStart()
            If rtde_connected Then
                Console.WriteLine("Beginning UR Control Start: " & name & vbNewLine)
            Else
                Console.WriteLine("Failed to begin UR Control Start: " & name & vbNewLine)
            End If
        Catch ex As Exception
            MsgBox("Error setting up the RTDE client")
        End Try
    End Sub

    'Opens up a network stream with the robot and allows for direct communication with strings
    Public Sub WriteTCP(listenFor As String, response As String)
        If simulate = True Then
            Exit Sub
        End If
        Try
            Dim tcpListener As New TcpListener(serverIPaddress, listeningPort)
            tcpListener.Start()
            Dim tcpClient As TcpClient = tcpListener.AcceptTcpClient()
            Dim stream As NetworkStream = tcpClient.GetStream()

            While tcpClient.Client.Connected()
                Dim arrayBytesRequest As Byte() = New Byte(tcpClient.Available - 1) {}
                Dim nRead = stream.Read(arrayBytesRequest, 0, arrayBytesRequest.Length)

                If nRead > 0 Then
                    Dim ReadStr = Encoding.ASCII.GetString(arrayBytesRequest)
                    If ReadStr = listenFor Then
                        Dim responseBytes = Encoding.ASCII.GetBytes(response & Environment.NewLine)
                        stream.Write(responseBytes, 0, responseBytes.Length)
                        tcpListener.Stop()
                        tcpClient.Close()
                        Exit While
                    Else
                        Console.WriteLine("Received something other than expected: " & ReadStr)
                    End If
                Else
                    If tcpClient.Available = 0 Then
                        stream.Close()
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox("Write TCP Error: " & ex.Message)
        End Try
    End Sub

    Public Function ReceiveTCP(listenFor As String)
        If simulate = True Then
            Return True
        End If
        Try
            Dim tcpListener As New TcpListener(serverIPaddress, listeningPort)
            tcpListener.Start()
            Dim tcpClient As TcpClient = tcpListener.AcceptTcpClient()
            Dim stream As NetworkStream = tcpClient.GetStream()

            While tcpClient.Client.Connected()
                Dim arrayBytesRequest As Byte() = New Byte(tcpClient.Available - 1) {}
                Dim nRead = stream.Read(arrayBytesRequest, 0, arrayBytesRequest.Length)

                If nRead > 0 Then
                    Dim ReadStr = Encoding.ASCII.GetString(arrayBytesRequest)
                    If ReadStr = listenFor Then
                        tcpListener.Stop()
                        tcpClient.Close()
                        Return True
                    Else
                        stream.Close()
                        Return False
                    End If
                Else
                    If tcpClient.Available = 0 Then
                        Return False
                    End If
                End If
            End While
            Return False
        Catch ex As Exception
            MsgBox("Receive TCP Error: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function SendAndReceive(cmd As String)
        Try
            If connected = True And simulate = False Then
                Dim cmdStr = cmd & Environment.NewLine
                Dim cmdBytes = Encoding.ASCII.GetBytes(cmdStr)
                robotSock.Send(cmdBytes)

                Threading.Thread.Sleep(100)
                Dim bytes(1024) As Byte
                Dim bytesRec As Integer = robotSock.Receive(bytes)
                If bytesRec <> 0 Then
                    Dim strRec = Encoding.ASCII.GetString(bytes, 0, bytesRec)
                    Return strRec
                Else
                    Return "Nothing"
                End If
            Else
                Return "Error"
            End If
        Catch ex As Exception
            MsgBox("Error when sending/receiving commands to the robot dashboard server: " & ex.Message)
            Return "Error"
        End Try
    End Function

    Public Function Play()
        Dim cmd = "play"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("Starting program") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Pause()
        Dim cmd = "pause"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("Pausing program") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function StopRobot()
        Dim cmd = "stop"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("Stopped") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Load(program As String)
        Dim cmd = "load " & program
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("Loading program") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Quit()
        Dim cmd = "quit"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("Disconnected") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Shutdown()
        Dim cmd = "shutdown"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("Shutting down") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Running()
        Dim cmd = "running"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("true") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckRobot()
        Dim cmd = "running"
        Dim response As String = SendAndReceive(cmd)
        If response = "Nothing" Or response = "Error" Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function GetLoadedProgram()
        Dim cmd = "get loaded program"
        Dim response As String = SendAndReceive(cmd)
        Return response.Split(": ")(1)
    End Function

    Public Function ProgramState()
        Dim cmd = "programState"
        Dim response As String = SendAndReceive(cmd)
        Return response
    End Function

    Public Function PowerOn()
        Dim cmd = "power on"
        Dim response As String = SendAndReceive(cmd)
        Return response
    End Function

    Public Function PowerOff()
        Dim cmd = "power off"
        Dim response As String = SendAndReceive(cmd)
        Return response
    End Function

    Public Function BrakeRelease()
        Dim cmd = "brake release"
        Dim response As String = SendAndReceive(cmd)
        Return response
    End Function

    Public Function SafetyStatus()
        Dim cmd = "safetystatus"
        Dim response As String = SendAndReceive(cmd)
        Return response.Split(": ")(1)
    End Function

    Public Function UnlockProtectiveStop()
        Dim cmd = "unlock protective stop"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("Protective stop releasing") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckRemoteMode()
        Dim cmd = "is in remote control"
        Dim response As String = SendAndReceive(cmd)
        If response.Contains("true") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function StartProgram(program As String)
        If Load(program) Then
            Threading.Thread.Sleep(250)
            If Play() Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    Public Function Robot_OnDataReceive()
        Return Nothing
    End Function

    Public Sub Robot_OnSockClosed()
        Console.WriteLine("Socket closed")
    End Sub
End Class
