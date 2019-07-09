Imports System.Runtime.InteropServices

Class Program

    Public Const SCARD_S_SUCCESS As UInteger = 0
    Public Const SCARD_W_REMOVED_CARD As UInteger = &H80100069UI
    Public Const SCARD_SCOPE_USER As UInteger = 0
    Public Const SCARD_AUTOALLOCATE As UInteger = &HFFFFFFFFUI
    Public Const SCARD_SHARE_SHARED As UInteger = &H2
    Public Const SCARD_PROTOCOL_T0 As UInteger = 1
    Public Const SCARD_PROTOCOL_T1 As UInteger = 2
    Public Const SCARD_LEAVE_CARD As UInteger = 0

    Public Class SCARD_IO_REQUEST
        Public dwProtocol As UInteger
        Public cbPciLength As UInteger
    End Class

    <DllImport("winscard.dll")>
    Public Shared Function SCardEstablishContext(dwScope As UInteger, pvReserved1 As IntPtr, pvReserved2 As IntPtr, ByRef phContext As IntPtr) As UInteger
    End Function

    <DllImport("winscard.dll")>
    Public Shared Function SCardListReaders(hContext As IntPtr, mszGroups() As Byte, mszReaders() As Byte, ByRef pcchReaders As UInteger) As UInteger
    End Function

    '<DllImport("winscard.dll")>
    'Public Shared Function SCardListReaders(hContext As IntPtr, ByRef mszGroups As String, ByRef mszReaders As String, ByRef pcchReaders As UInteger) As UInteger
    'End Function

    <DllImport("winscard.dll")>
    Public Shared Function SCardConnect(hContext As IntPtr, szReader As String, dwShareMode As UInteger, dwPreferredProtocols As UInteger, ByRef phCard As IntPtr, ByRef pdwActiveProtocol As UInteger) As UInteger
    End Function

    <DllImport("winscard.dll")>
    Public Shared Function SCardTransmit(hCard As IntPtr, pioSendPci As IntPtr, pbSendBuffer As Byte(), cbSendLength As UInteger, pioRecvPci As SCARD_IO_REQUEST, pbRecvBuffer As Byte(), ByRef pcbRecvLength As UInteger) As UInteger
    End Function

    <DllImport("winscard.dll")>
    Public Shared Function SCardDisconnect(hCard As IntPtr, dwDisposition As UInteger) As UInteger
    End Function

    <DllImport("winscard.dll")>
    Public Shared Function SCardReleaseContext(hContext As IntPtr) As UInteger
    End Function

    <DllImport("winscard.dll")>
    Public Shared Function SCardFreeMemory(hContext As IntPtr, pvMem As IntPtr) As UInteger
    End Function

    <DllImport("winscard.dll")>
    Public Shared Function SCardFreeMemory(hContext As IntPtr, pvMem As String) As UInteger
    End Function

    Public Shared Sub Main(args() As String)

        Console.WriteLine("start")

        Dim hContext As IntPtr
        Dim ret = SCardEstablishContext(SCARD_SCOPE_USER, IntPtr.Zero, IntPtr.Zero, hContext)
        If ret <> SCARD_S_SUCCESS Then

            Console.WriteLine($"error SCardEstablishContext {ret}")
            Return
        End If

        Try
            'Dim pcchReaders = SCARD_AUTOALLOCATE
            'Dim mszReaders As String = Nothing
            'ret = SCardListReaders(hContext, Nothing, mszReaders, pcchReaders)
            'If ret <> SCARD_S_SUCCESS Then

            '    Console.WriteLine($"error SCardListReaders {ret}")
            '    Return
            'End If

            Dim pcchReaders As UInteger
            ret = SCardListReaders(hContext, Nothing, Nothing, pcchReaders)
            If ret <> SCARD_S_SUCCESS Then

                Console.WriteLine($"error SCardListReaders {ret}")
                Return
            End If

            Dim mszReaders = New Byte(CInt(pcchReaders)) {}
            ret = SCardListReaders(hContext, Nothing, mszReaders, pcchReaders)
            If ret <> SCARD_S_SUCCESS Then

                Console.WriteLine($"error SCardListReaders {ret}")
                Return
            End If
            Dim reader = Text.Encoding.UTF8.GetString(mszReaders)

            Try
                Console.WriteLine($"ok SCardListReaders {reader}")

                Dim hCard As IntPtr
                Dim dwActiveProtocol As UInt32
                Console.WriteLine($"polling SCardConnect {ret}")
                Do While True

                    ret = SCardConnect(hContext, reader, SCARD_SHARE_SHARED, SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1, hCard, dwActiveProtocol)
                    If ret <> SCARD_W_REMOVED_CARD Then Exit Do

                    Threading.Thread.Sleep(100)
                Loop
                If ret <> SCARD_S_SUCCESS Then

                    Console.WriteLine($"error SCardConnect {ret}")
                    Return
                End If

                Try
                    Console.WriteLine($"ok SCardConnect {dwActiveProtocol}")

                    Dim result = SCardTransmit(hCard, {&H0, &HA4, &H0, &H0})
                    'WriteLine("SelectMF:", result)

                    For Each aid In {
                            New Byte() {&H0, &HA4, &H4, &H0, &H7, &HA0, &H0, &H0, &H0, &H65, &H10, &H10, &H0},  ' JCB    A0-00-00-00-65-10-10
                            New Byte() {&H0, &HA4, &H4, &H0, &H7, &HA0, &H0, &H0, &H0, &H3, &H10, &H10, &H0},   ' VISA   A0-00-00-00-03-10-10
                            New Byte() {&H0, &HA4, &H4, &H0, &H7, &HA0, &H0, &H0, &H0, &H4, &H10, &H10, &H0},   ' MASTER A0-00-00-00-04-10-10
                            New Byte() {&H0, &HA4, &H4, &H0, &H7, &HA0, &H0, &H0, &H1, &H52, &H10, &H10, &H0},  ' Diners A0-00-00-01-52-10-10
                            New Byte() {&H0, &HA4, &H4, &H0, &H5, &HA0, &H0, &H0, &H0, &H25, &H0},              ' AMEX   A0-00-00-00-25-xx-xx
                            New Byte() {&H0, &HA4, &H4, &H0, &H3, &HA0, &H0, &H0, &H0}                          ' unknown
                        }

                        result = SCardTransmit(hCard, aid)
                        If result(result.Count - 2) <> &H90 OrElse result(result.Count - 1) <> &H0 Then Continue For

                        'WriteLine("SendCommand:", result)
                        Console.WriteLine($"App DF: {ConvertHex(result, 4, result(3))}")

                        result = SCardTransmit(hCard, {&H0, &HB2, &H1, &HC, &H0})
                        WriteLine("SendCommand:", result)

                        Exit For
                    Next

                    'Console.WriteLine(Text.Encoding.ASCII.GetString(result))

                    ' ToDo: 17桁以上のカードはどうなるの？
                    ' 「D」が終端記号
                    Dim bit4s = Split4bit(result)
                    Dim sep = IndexOf(bit4s, &HD, 8)
                    Dim cardno = ConvertNum(SubArray(bit4s, 8, sep - 8), 0, sep - 8)

                    ' カード番号の終端記号から「Dy ym mx」のフォーマットで入っている
                    Dim year = CInt(bit4s(sep + 1) * 10 + bit4s(sep + 2))
                    Dim month = CInt(bit4s(sep + 3) * 10 + bit4s(sep + 4))

                    ' ToDo: 26～52バイト目の位置に名前が入っていると仮定する、右スペース埋め、「/」区切り
                    ' ToDo: 0x90が終端記号か？
                    ' ToDo: AMEXが入ってないので取得しない(2:5F20には入っている模様)
                    'Dim name = Text.Encoding.ASCII.GetString(result, 26, 26).TrimEnd

                    Console.WriteLine($"CardNo: {cardno}")
                    Console.WriteLine($"YM    : {month:00}/{year:00}")
                    'Console.WriteLine($"Name  : {name}")


                Finally
                    SCardDisconnect(hCard, SCARD_LEAVE_CARD)

                End Try

            Finally
                'SCardFreeMemory(hContext, mszReaders)

            End Try

        Catch ex As Exception
            Console.WriteLine($"error {ex.Message}")

        Finally

            SCardReleaseContext(hContext)
        End Try
        Console.WriteLine("end")
        Console.WriteLine()
        Console.WriteLine("push any key...")
        Console.ReadKey()
    End Sub

    Public Shared Function SCardTransmit(hCard As IntPtr, send() As Byte) As Byte()

        Dim SCARD_PCI_T1 As IntPtr
        Dim recv = New Byte(261) {}
        Dim recv_len = CUInt(recv.Length)

        'WriteLine("SCardTransmit(Send):", send)
        Dim ret = SCardTransmit(hCard, SCARD_PCI_T1, send, CUInt(send.Length), Nothing, recv, recv_len)
        If ret <> SCARD_S_SUCCESS Then

            Throw New Exception($"error SCardTransmit {ret}")
        End If

        Dim result = New Byte(CInt(recv_len - 1)) {}
        Array.Copy(recv, result, recv_len)
        'WriteLine("SCardTransmit(Recv):", result)
        Return result
    End Function

    Public Shared Sub WriteLine(title As String, xs() As Byte)

        Console.WriteLine($"{title} {ConvertHex(xs, 0, xs.Count)}")
    End Sub

    Public Shared Function ConvertHex(xs() As Byte, start As Integer, count As Integer) As String

        If count = 0 Then Return ""

        Dim s = ""
        For i = start To start + count - 1

            s += $" {xs(i):X2}"
        Next

        Return s.Substring(1)
    End Function

    Public Shared Function ConvertNum(xs() As Byte, start As Integer, count As Integer) As String

        If count = 0 Then Return ""

        Dim s = ""
        For i = start To start + count - 1

            s += $"{xs(i):X}"
        Next

        Return s
    End Function

    Public Shared Function Split4bit(xs() As Byte) As Byte()

        Dim ns = New Byte(xs.Count * 2 - 1) {}
        For i = 0 To xs.Count - 1

            ns(i * 2 + 0) = xs(i) >> 4
            ns(i * 2 + 1) = CByte(xs(i) And &HF)
        Next

        Return ns
    End Function

    Public Shared Function IndexOf(Of T)(xs() As T, find As T, start As Integer) As Integer

        For i = start To xs.Count - 1

            If Object.Equals(xs(i), find) Then Return i
        Next
        Return -1
    End Function

    Public Shared Function SubArray(Of T)(xs() As T, start As Integer, count As Integer) As T()

        Return New List(Of T)(xs).Skip(start).Take(count).ToArray
    End Function
End Class
