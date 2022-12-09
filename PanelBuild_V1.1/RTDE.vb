Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Net.Sockets
Imports System.Threading
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace Ur_Rtde
    Interface IEncoderDecoder
        Function Decode(ByRef o As Object, ByVal buf As Byte(), ByRef offset As Integer) As Object
        Sub Encode(ByVal o As Object, ByVal buf As Byte(), ByRef offset As Integer)
    End Interface

    Public Class EncodeValue
        Implements IEncoderDecoder

        Private type As Type
        Private Typesize As Integer

        Public Sub New(ByVal type As Type)
            Me.type = type
            Typesize = Marshal.SizeOf(type)
            If type.FullName = "System.Boolean" Then
                Typesize = 1
            End If
        End Sub

        Public Sub IEncoderDecoder_Encode(ByVal o As Object, ByVal buf As Byte(), ByRef offset As Integer) Implements IEncoderDecoder.Encode
            Dim b As Byte() = Nothing

            Select Case type.FullName
                Case "System.Boolean"
                    b = BitConverter.GetBytes(CBool(o))
                Case "System.Byte"
                    b = New Byte(1) {}
                    b(0) = CByte(o)
                Case "System.UInt32"
                    b = BitConverter.GetBytes(CUInt(o))
                Case "System.Int32"
                    b = BitConverter.GetBytes(CInt(o))
                Case "System.UInt64"
                    b = BitConverter.GetBytes(CULng(o))
                Case "System.Double"
                    b = BitConverter.GetBytes(CDbl(o))
            End Select

            If BitConverter.IsLittleEndian Then Array.Reverse(b)
            Array.Copy(b, 0, buf, offset, Typesize)
            offset += Typesize
        End Sub

        Public Function IEncoderDecoder_Decode(ByRef o As Object, ByVal buf As Byte(), ByRef offset As Integer) As Object Implements IEncoderDecoder.Decode
            Dim b As Byte() = New Byte(Typesize - 1) {}
            Array.Copy(buf, offset, b, 0, Typesize)
            If BitConverter.IsLittleEndian Then Array.Reverse(b)
            offset += Typesize

            Select Case type.FullName
                Case "System.Boolean"
                    Return BitConverter.ToBoolean(b, 0)
                Case "System.Byte"
                    Return b(0)
                Case "System.UInt32"
                    Return BitConverter.ToUInt32(b, 0)
                Case "System.Int32"
                    Return BitConverter.ToInt32(b, 0)
                Case "System.UInt64"
                    Return BitConverter.ToUInt64(b, 0)
                Case "System.Double"
                    Return BitConverter.ToDouble(b, 0)
            End Select

            Return Nothing
        End Function
    End Class

    Public Class EncodeArray
        Implements IEncoderDecoder

        Private ArraySize, Typesize As Integer
        Private type As Type

        Public Sub New(ByVal size As Integer, ByVal type As Type)
            ArraySize = size
            Typesize = Marshal.SizeOf(type)
            Me.type = type
        End Sub

        Public Sub IEncoderDecoder_Encode(ByVal o As Object, ByVal buf As Byte(), ByRef offset As Integer) Implements IEncoderDecoder.Encode
            Dim array As Array = TryCast(o, Array)

            For i As Integer = 0 To ArraySize - 1
                Dim b As Byte() = Nothing

                Select Case type.FullName
                    Case "System.UInt32"
                        b = BitConverter.GetBytes(CUInt(array.GetValue(i)))
                    Case "System.Int32"
                        b = BitConverter.GetBytes(CInt(array.GetValue(i)))
                    Case "System.UInt64"
                        b = BitConverter.GetBytes(CULng(array.GetValue(i)))
                    Case "System.Double"
                        b = BitConverter.GetBytes(CDbl(array.GetValue(i)))
                End Select

                If BitConverter.IsLittleEndian Then Array.Reverse(b)
                Array.Copy(b, 0, buf, offset, Typesize)
                offset += Typesize
            Next
        End Sub

        Public Function IEncoderDecoder_Decode(ByRef o As Object, ByVal buf As Byte(), ByRef offset As Integer) As Object Implements IEncoderDecoder.Decode
            Dim obj As Array = TryCast(o, Array)

            For i As Integer = 0 To ArraySize - 1
                Dim b As Byte() = New Byte(Typesize - 1) {}
                Array.Copy(buf, offset, b, 0, Typesize)
                If BitConverter.IsLittleEndian Then Array.Reverse(b)
                offset += Typesize
                Dim value As Object = Nothing

                Select Case type.FullName
                    Case "System.UInt32"
                        value = BitConverter.ToUInt32(b, 0)
                    Case "System.Int32"
                        value = BitConverter.ToInt32(b, 0)
                    Case "System.UInt64"
                        value = BitConverter.ToUInt64(b, 0)
                    Case "System.Double"
                        value = BitConverter.ToDouble(b, 0)
                End Select

                obj.SetValue(value, i)
            Next

            Return obj
        End Function
    End Class

    Public Class RtdeClient
        Enum RTDE_Command
            REQUEST_PROTOCOL_VERSION = 86
            GET_URCONTROL_VERSION = 118
            TEXT_MESSAGE = 77
            DATA_PACKAGE = 85
            CONTROL_PACKAGE_SETUP_OUTPUTS = 79
            CONTROL_PACKAGE_SETUP_INPUTS = 73
            CONTROL_PACKAGE_START = 83
            CONTROL_PACKAGE_PAUSE = 80
        End Enum

        Public TimeOut As Integer = 500
        Public sock As TcpClient = New TcpClient()
        Private receiveDone As ManualResetEvent = New ManualResetEvent(False)
        Public Property ProtocolVersion As UInteger
        Public bufRecv As Byte() = New Byte(1500) {}

        Public Event OnDataReceive As EventHandler
        Public Event OnSockClosed As EventHandler

        Public Outputs_Recipe_Id, Inputs_Recipe_Id As Byte
        Public UrStructOuput, UrStructInput As Object

        Private UrStructOuputDecoder, UrStructInputDecoder As IEncoderDecoder()
        Public Property ErrorMessage As String

        Public ControlConnected As Boolean
        Protected disposed As Boolean = False

        Public Function Connect(ByVal host As String, ByVal Optional ProtocolVersion As UInteger = 2, ByVal Optional timeOut As Integer = 500) As Boolean
            Dim InternalbufRecv As Byte() = New Byte(bufRecv.Length) {}
            Me.ProtocolVersion = 1

            Try
                sock.Connect(host, 30004)
                sock.Client.BeginReceive(InternalbufRecv, 0, InternalbufRecv.Length, SocketFlags.None, AddressOf AsynchReceive, InternalbufRecv)
                If ProtocolVersion <> 1 Then Set_UR_Protocol_Version(ProtocolVersion)
                disposed = False
                ControlConnected = True
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Protected Overrides Sub Finalize()
            If disposed = False Then
                Disconnect(False)
            End If
        End Sub
        Public Sub Disconnect(ByVal boolReuse As Boolean)
            Try
                If disposed = False Then
                    Ur_ControlPause()
                    Dim InternalBufRecv As Byte() = New Byte(bufRecv.Length) {}
                    If boolReuse Then
                        sock.Client.BeginDisconnect(True, AddressOf AsynchEnd, InternalBufRecv)
                    Else
                        sock.Client.BeginDisconnect(False, AddressOf AsynchEnd, InternalBufRecv)
                    End If
                    disposed = True
                    ControlConnected = False
                End If
            Catch ex As Exception
                Console.WriteLine("Error disconnecting RTDE client")
            End Try

        End Sub

        Private Sub AsynchEnd(ByVal ar As IAsyncResult)
            'Do nothing
        End Sub

        Private Sub AsynchReceive(ByVal ar As IAsyncResult)
            Dim bytesRead As Integer

            bytesRead = sock.Client.EndReceive(ar)

            Dim InternalbufRecv As Byte() = CType(ar.AsyncState, Byte())

            If bytesRead > 0 Then

                SyncLock bufRecv
                    Array.Copy(InternalbufRecv, bufRecv, InternalbufRecv.Length - 1)
                End SyncLock

                If InternalbufRecv(2) = CByte(RTDE_Command.TEXT_MESSAGE) Then

                    If ProtocolVersion = 1 Then
                        ErrorMessage = Encoding.ASCII.GetString(InternalbufRecv, 4, InternalbufRecv(1) - 4 - 2)
                    Else
                        ErrorMessage = Encoding.ASCII.GetString(InternalbufRecv, 4, InternalbufRecv(3))
                    End If
                End If

                receiveDone.[Set]()
                sock.Client.BeginReceive(InternalbufRecv, 0, InternalbufRecv.Length, SocketFlags.None, AddressOf AsynchReceive, InternalbufRecv)

                Try

                    If bufRecv(2) = CByte(RTDE_Command.DATA_PACKAGE) Then
                        Dim offset As Integer = 3

                        If ProtocolVersion = 2 Then
                            offset += 1
                            If bufRecv(3) <> Outputs_Recipe_Id Then Return
                        End If

                        Dim f As FieldInfo() = UrStructOuput.[GetType]().GetFields()

                        For i As Integer = 0 To f.Length - 1
                            Dim currentvalue As Object = f(i).GetValue(UrStructOuput)
                            If f(i).FieldType.IsArray Then
                                UrStructOuputDecoder(i).Decode(currentvalue, bufRecv, offset)
                            Else
                                f(i).SetValue(UrStructOuput, UrStructOuputDecoder(i).Decode(currentvalue, bufRecv, offset))
                            End If
                        Next

                        RaiseEvent OnDataReceive(Me, Nothing)
                    End If

                Catch
                End Try
            Else
                Try
                    RaiseEvent OnSockClosed(Me, Nothing)
                Catch ex As Exception
                End Try
            End If
        End Sub

        Private Sub SendRtdePacket(ByVal RTDEType As RTDE_Command, ByVal Optional payload As Byte() = Nothing)
            ErrorMessage = Nothing
            Try
                If payload Is Nothing Then payload = New Byte(-1) {}
                Dim s As Byte() = New Byte(payload.Length + 3 - 1) {}
                Dim size As Byte() = BitConverter.GetBytes(payload.Length + 3)
                s(0) = size(1)
                s(1) = size(0)
                s(2) = CByte(RTDEType)
                If payload IsNot Nothing Then Array.Copy(payload, 0, s, 3, payload.Length)
                receiveDone.Reset()
                sock.Client.BeginSend(s, 0, s.Length, SocketFlags.None, Nothing, Nothing)
            Catch ex As Exception
                Console.WriteLine("Issue sending RTDE Packet")
            End Try

        End Sub

        Private Function Send_UR_Command(ByVal Cmd As RTDE_Command, ByVal Optional payload As Byte() = Nothing) As Boolean
            SendRtdePacket(Cmd, payload)

            If receiveDone.WaitOne(TimeOut) Then

                SyncLock bufRecv
                    Return (bufRecv(2) = CByte(Cmd))
                End SyncLock
            End If

            Return False
        End Function

        Private Function Set_UR_Protocol_Version(ByVal Version As UInteger) As Boolean
            Dim V As Byte() = {0, CByte(Version)}
            Dim ret As Boolean = Send_UR_Command(RTDE_Command.REQUEST_PROTOCOL_VERSION, V)
            If (ret = True) AndAlso (bufRecv(3) = 1) Then ProtocolVersion = Version
            Return ret
        End Function

        Public Function Ur_ControlStart() As Boolean
            ControlConnected = Send_UR_Command(RTDE_Command.CONTROL_PACKAGE_START)
            Return ControlConnected
        End Function

        Public Function Ur_ControlPause() As Boolean
            Return Send_UR_Command(RTDE_Command.CONTROL_PACKAGE_PAUSE)
        End Function

        Public Function Send_Ur_Inputs() As Boolean
            Dim f As FieldInfo() = UrStructInput.[GetType]().GetFields()
            Dim buf As Byte() = New Byte(1499) {}
            Dim offset As Integer = 0

            For i As Integer = 0 To f.Length - 1
                UrStructInputDecoder(i).Encode(f(i).GetValue(UrStructInput), buf, offset)
            Next

            Dim payload As Byte()

            If ProtocolVersion = 1 Then
                payload = New Byte(offset - 1) {}
                Array.Copy(buf, payload, offset)
            Else
                payload = New Byte(offset + 1 - 1) {}
                payload(0) = Inputs_Recipe_Id
                Array.Copy(buf, 0, payload, 1, offset)
            End If

            Send_UR_Command(RTDE_Command.DATA_PACKAGE, payload)
            Return True
        End Function

        Private Function Setup_Ur_InputsOutputs(ByVal Cmd As RTDE_Command, ByVal UrStruct As Object, <Out> ByRef encoder As IEncoderDecoder(), ByVal Optional Frequency As Double = 1) As Boolean
            Dim f As FieldInfo() = UrStruct.[GetType]().GetFields()
            encoder = New IEncoderDecoder(f.Length - 1) {}
            Dim b As StringBuilder = New StringBuilder()

            For i As Integer = 0 To f.Length - 1
                b.Append((If(i = 0, "", ",")) & f(i).Name)

                If f(i).FieldType.IsArray Then
                    Dim array As Array = TryCast(f(i).GetValue(UrStruct), Array)
                    Dim element As Object = array.GetValue(0)
                    encoder(i) = New EncodeArray(array.Length, element.[GetType]())
                Else
                    encoder(i) = New EncodeValue(f(i).FieldType)
                End If
            Next

            Dim payload As Byte()

            If (Cmd = RTDE_Command.CONTROL_PACKAGE_SETUP_OUTPUTS) AndAlso (ProtocolVersion = 2) Then
                payload = New Byte(b.Length + 8 - 1) {}
                Dim Freq As Byte() = BitConverter.GetBytes(Frequency)
                If BitConverter.IsLittleEndian Then Array.Reverse(Freq)
                Array.Copy(Freq, 0, payload, 0, 8)
                Array.Copy(Encoding.ASCII.GetBytes(b.ToString()), 0, payload, 8, b.Length)
            Else
                payload = Encoding.ASCII.GetBytes(b.ToString())
            End If

            If Send_UR_Command(Cmd, payload) = True Then

                If Cmd = RTDE_Command.CONTROL_PACKAGE_SETUP_OUTPUTS Then
                    Outputs_Recipe_Id = bufRecv(3)
                Else
                    Inputs_Recipe_Id = bufRecv(3)
                End If

                Dim s As String = Encoding.ASCII.GetString(bufRecv, 3, bufRecv.Length - 3)
                If s.Contains("NOT_FOUND") Then Return False
                If s.Contains("IN_USE") Then Return False
                Return True
            End If

            Return False
        End Function

        Public Function Setup_Ur_Outputs(ByVal UrStruct As Object, ByVal Optional Frequency As Double = 1) As Boolean
            Me.UrStructOuput = UrStruct
            Return Setup_Ur_InputsOutputs(RTDE_Command.CONTROL_PACKAGE_SETUP_OUTPUTS, UrStruct, UrStructOuputDecoder, Frequency)
        End Function

        Public Function Setup_Ur_Inputs(ByVal UrStruct As Object) As Boolean
            Me.UrStructInput = UrStruct
            Return Setup_Ur_InputsOutputs(RTDE_Command.CONTROL_PACKAGE_SETUP_INPUTS, UrStruct, UrStructInputDecoder)
        End Function
    End Class
End Namespace