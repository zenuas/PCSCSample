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
                Do While True

                    ret = SCardConnect(hContext, reader, SCARD_SHARE_SHARED, SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1, hCard, dwActiveProtocol)
                    If ret <> SCARD_W_REMOVED_CARD Then Exit Do

                    Console.WriteLine($"polling SCardConnect {ret}")
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

                    result = SCardTransmit(hCard, {&H0, &HA4, &H4, &H0, &H7, &HA0, &H0, &H0, &H0, &H4, &H10, &H10, &H0})
                    'WriteLine("SendCommand:", result)

                    result = SCardTransmit(hCard, {&H0, &HB2, &H1, &HC, &H0})
                    'WriteLine("SendCommand:", result)

                    'Console.WriteLine(Text.Encoding.ASCII.GetString(result))

                    ' ToDo: AMEXとかは磁気カードと同じ前ゼロ埋めか？
                    ' ToDo: 17桁以上のカードはどうなるの？「D」が終端記号か？
                    Dim cardno = ""
                    For i = 4 To 11

                        cardno += result(i).ToString("X2")
                        If i < 11 AndAlso i Mod 2 = 1 Then cardno += "-"
                    Next

                    ' ToDo: 12バイト目からの位置に「Dy ym mx」のフォーマットで入っていると仮定する
                    Dim ym = result(12).ToString("X2") + result(13).ToString("X2") + result(14).ToString("X2")
                    Dim year = CInt(ym.Substring(1, 2))
                    Dim month = CInt(ym.Substring(3, 2))

                    ' ToDo: 26～52バイト目の位置に名前が入っていると仮定する、右スペース埋め、「/」区切り
                    ' ToDo: 0x90が終端記号か？
                    Dim name = Text.Encoding.ASCII.GetString(result, 26, 26).TrimEnd

                    Console.WriteLine($"CardNo: {cardno}")
                    Console.WriteLine($"YM    : {month:00}/{year:00}")
                    Console.WriteLine($"Name  : {name}")


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

        Dim ret = SCardTransmit(hCard, SCARD_PCI_T1, send, CUInt(send.Length), Nothing, recv, recv_len)
        If ret <> SCARD_S_SUCCESS Then

            Throw New Exception($"error SCardTransmit {ret}")
        End If

        Dim result = New Byte(CInt(recv_len - 1)) {}
        Array.Copy(recv, result, recv_len)
        Return result
    End Function

    Public Shared Sub WriteLine(title As String, xs() As Byte)

        Console.Write(title)
        For i = 0 To xs.Length - 1

            Console.Write($" {xs(i):X2}")
        Next
        Console.WriteLine()
    End Sub
End Class
