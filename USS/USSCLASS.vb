Imports System.IO.Ports
'Author: Mostafa Soghandi
'Email: Mostafa.soghandi@gmail.com
Public Class USSCLASS
    Public Event RData(ByVal sender As Object, ByVal e As EventArgs)

    Dim MSComm1 As New SerialPort("COM1", 19200, Parity.Even, 8, StopBits.One)

    Public Sub PortOpen()
        MSComm1.ReadBufferSize = 512
        MSComm1.WriteBufferSize = 512
        MSComm1.ReceivedBytesThreshold = 15
        MSComm1.DtrEnable = True
        MSComm1.RtsEnable = True
        AddHandler MSComm1.DataReceived, AddressOf RecieveData
        MSComm1.Open()
    End Sub
    Public RcvStr As String = ""
    Public Function RecieveData() As Boolean
        Dim i%, buf
        Dim hexdisp As String = Nothing
        Dim inByte(127) As Byte
        buf = ""
        MSComm1.Read(inByte, 0, 127)

        Dim jm As Integer = inByte(1).ToString()
        For i = 0 To jm - 1
            buf = buf + IIf(inByte(i) > 15, Hex(inByte(i)), "0" + Hex(inByte(i))) + Chr(32)
        Next i
        hexdisp = hexdisp + buf
        RcvStr = jm.ToString + ":" + hexdisp.ToString
        RaiseEvent RData(Me, New System.EventArgs)
    End Function

    Public Sub PortClose()
        MSComm1.Close()
    End Sub

    Public Function RunMMS(ByVal Freq As Integer, ByVal Address As Integer) As Boolean
        Dim i(15) As Byte
        Dim j As Integer
        Dim Pin As Single
        Dim PinH As String = ""
        Dim PinL As String = ""
        Pin = Freq * 16384 / 50

        If Len(Hex(Pin)) = 4 Then
            PinH = Mid(Hex(Pin), 1, 2)
            PinL = Mid(Hex(Pin), 3, 2)
        End If

        If Len(Hex(Pin)) < 4 Then
            PinH = Mid(Hex(Pin), 1, 1)
            PinL = Mid(Hex(Pin), 2, 2)
        End If

        i(0) = &H2
        i(1) = &HE
        i(2) = "&H" & Hex(Address).ToString
        MsgBox(i(2))
        i(3) = &H0
        i(4) = &H0
        i(5) = &H0
        i(6) = &H0
        i(7) = &H0
        i(8) = &H0
        i(9) = &H0
        i(10) = &H0

        i(11) = &H4
        i(12) = &H7F
        i(13) = "&H" & PinH.ToString
        i(14) = "&H" & PinL.ToString

        For j = 0 To 14
            i(15) = i(15) Xor i(j)
        Next j
        MSComm1.Write(i, 0, i.Length)
    End Function

    Public Function StopRunning(ByVal Address As Integer) As Boolean
        Dim i(15) As Byte
        Dim j As Integer
        i(0) = &H2
        i(1) = &HE
        i(2) = "&H" & Hex(Address).ToString

        i(3) = &H0
        i(4) = &H0
        i(5) = &H0
        i(6) = &H0
        i(7) = &H0
        i(8) = &H0
        i(9) = &H0
        i(10) = &H0

        i(11) = &H4
        i(12) = &H7E
        i(13) = &H0
        i(14) = &H0

        For j = 0 To 14
            i(15) = i(15) Xor i(j)
        Next j
        MSComm1.Write(i, 0, i.Length)
    End Function

    Public Function ReverseRun(ByVal Freq As Integer, ByVal Address As Integer) As Boolean
        Dim i(15) As Byte
        Dim j As Integer
        Dim Pin As Single
        Dim PinH As String = ""
        Dim PinL As String = ""
        Pin = Val(Freq) * 16384 / 50

        If Len(Hex(Pin)) = 4 Then
            PinH = Mid(Hex(Pin), 1, 2)
            PinL = Mid(Hex(Pin), 3, 2)
        End If

        If Len(Hex(Pin)) < 4 Then
            PinH = Mid(Hex(Pin), 1, 1)
            PinL = Mid(Hex(Pin), 2, 2)
        End If

        i(0) = &H2
        i(1) = &HE
        i(2) = "&H" & Hex(Address).ToString
        i(3) = &H0
        i(4) = &H0
        i(5) = &H0
        i(6) = &H0
        i(7) = &H0
        i(8) = &H0
        i(9) = &H0
        i(10) = &H0

        i(11) = &HC
        i(12) = &H7F
        i(13) = "&H" + PinH.ToString
        i(14) = "&H" + PinL.ToString

        For j = 0 To 14
            i(15) = i(15) Xor i(j)
        Next j
        MSComm1.Write(i, 0, i.Length)
    End Function

    Public Function RunMMSOnJOGMode(ByVal Address As Integer) As Boolean
        Dim i(15) As Byte
        Dim j As Integer

        i(0) = &H2
        i(1) = &HE
        i(2) = "&H" & Hex(Address).ToString

        i(3) = &H0
        i(4) = &H0
        i(5) = &H0
        i(6) = &H0
        i(7) = &H0
        i(8) = &H0
        i(9) = &H0
        i(10) = &H0

        i(11) = &H5
        i(12) = &H7E
        i(13) = &H0
        i(14) = &H0


        For j = 0 To 14
            i(15) = i(15) Xor i(j)
        Next j
        MSComm1.Write(i, 0, i.Length)

    End Function

    Public Function ReverseJOG(ByVal address As Integer) As Boolean
        Dim i(15) As Byte
        Dim j As Integer
        i(0) = &H2
        i(1) = &HE
        i(2) = "&H" & Hex(address).ToString

        i(3) = &H0
        i(4) = &H0
        i(5) = &H0
        i(6) = &H0
        i(7) = &H0
        i(8) = &H0
        i(9) = &H0
        i(10) = &H0

        i(11) = &H6
        i(12) = &H7E
        i(13) = &H0
        i(14) = &H0

        For j = 0 To 14
            i(15) = i(15) Xor i(j)
        Next j
        MSComm1.Write(i, 0, i.Length)
    End Function
    Public Function StopJOG(ByVal address As Integer) As Boolean
        Dim i(15) As Byte
        Dim j As Integer
        i(0) = &H2
        i(1) = &HE
        i(2) = "&H" & Hex(address).ToString

        i(3) = &H0
        i(4) = &H0
        i(5) = &H0
        i(6) = &H0
        i(7) = &H0
        i(8) = &H0
        i(9) = &H0
        i(10) = &H0

        i(11) = &H4
        i(12) = &H7E
        i(13) = &H0
        i(14) = &H0

        For j = 0 To 14
            i(15) = i(15) Xor i(j)
        Next j
        MSComm1.Write(i, 0, i.Length)
    End Function

    Public Function ReqParam(ByVal address As Integer) As Boolean
        Dim i(15) As Byte
        Dim j As Integer
        i(0) = &H2
        i(1) = &HE
        i(2) = "&H" & Hex(address).ToString

        i(3) = &H10
        i(4) = &H3
        i(5) = &H0
        i(6) = &H0
        i(7) = &H0
        i(8) = &H3
        i(9) = &H0
        i(10) = &H0

        i(11) = &H0
        i(12) = &H0
        i(13) = &H0
        i(14) = &H0

        For j = 0 To 14
            i(15) = i(15) Xor i(j)
        Next j

        MSComm1.Write(i, 0, i.Length)
    End Function

End Class
