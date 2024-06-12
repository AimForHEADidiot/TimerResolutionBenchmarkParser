Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.Globalization

Module Program

    Const qChr As Char = Chr(34)
    Private fileText As String
    Private currentResult As v3Result = New v3Result
    Private lowestResult As v3Result = New v3Result(1009.0F, 1009.0F, 1009.0F)

    <DllImport("user32.dll")>
    Private Function OpenClipboard(ByVal hWndNewOwner As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function CloseClipboard() As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function SetClipboardData(ByVal uFormat As UInteger, ByVal data As IntPtr) As Boolean
    End Function

    Structure v3Result
        Public RequestedResolution As Single
        Public Delta As Single
        Public STDEV As Single

        Sub New(Optional _RequestedResolution As Single = 1337.0F, Optional _Delta As Single = 1337.0F, Optional _STDEV As Single = 1337.0F)
            RequestedResolution = _RequestedResolution
            Delta = _Delta
            STDEV = _STDEV
        End Sub

        Public Sub Become(ByVal Target As v3Result)
            RequestedResolution = Target.RequestedResolution
            Delta = Target.Delta
            STDEV = Target.STDEV
        End Sub

        Public Function StringToStructure(ByVal Source As String) As v3Result
            If Regex.Split(Source, ",").Length - 1 >= 2 Then
                Dim temp = Source.Split(",")
                Return New v3Result(Convert.ToSingle(temp(0), CultureInfo.InvariantCulture), Convert.ToSingle(temp(1), CultureInfo.InvariantCulture), Convert.ToSingle(temp(2), CultureInfo.InvariantCulture))
            End If
            Return New v3Result(1337.0F, 1337.0F, 1337.0F)
        End Function

        Public Function ToArgLine() As String
            Dim rrArg As String = ""
            Dim rr() As Char = RequestedResolution.ToString.ToCharArray
            For i As Integer = 2 To rr.Length - 1
                rrArg += rr(i)
            Next
            If rrArg.Length < 4 Then
                Select Case rrArg.Length
                    Case 1
                        rrArg += "000"
                    Case 2
                        rrArg += "00"
                    Case 3
                        rrArg += "0"
                End Select
            End If
            Return ("-resolution " & rrArg & " -no-console")
        End Function

        Public Overrides Function ToString() As String
            Return ("Requested resolution ms: " & RequestedResolution.ToString & " │ Delta ms: " & Delta.ToString & " │ STDEV ms: " & STDEV.ToString)
        End Function
    End Structure

    Sub Main(args As String())
        Console.Title = "ResultsParser by HEAD+ angledemon"
        Console.WriteLine("ResultsParser by HEAD+ angledemon")
        Console.WriteLine("------------------------------------")
        Console.WriteLine("Press any key to start parser . . . ")
        Console.ReadKey(True)
        If Not IO.File.Exists("results.txt") Then
            Console.WriteLine("File " & qChr & "results.txt" & qChr & " does not exist or isn't placed in same folder as parser.")
            Console.WriteLine("Press any key to close parser . . . ")
            Console.ReadKey(True)
        Else
            Console.WriteLine("File " & qChr & "results.txt" & qChr & " has been found.")
            Console.WriteLine("Reading file . . .")
            fileText = ReadFile()
            Console.WriteLine("Starting to parse file . . .")
            ParseResults(fileText.Split(vbCrLf))
            Console.WriteLine("Copy shortcut arguments to clipboard?")
            Console.WriteLine("Please answer Y/N to continue . . .")
            Dim xKey = Console.ReadKey
            If xKey.Key = ConsoleKey.Y Then ToClipboard(lowestResult.ToArgLine)
            If xKey.Key = ConsoleKey.N Then
                Console.WriteLine("")
                Console.WriteLine(lowestResult.ToArgLine)
            End If
            Console.WriteLine("Press any key to close parser . . .")
            Console.ReadKey(True)
        End If
    End Sub
    Sub ParseResults(ByVal StringLines() As String)
        For i As Integer = 1 To StringLines.Length - 2
            currentResult = currentResult.StringToStructure(StringLines(i))
            Console.WriteLine(currentResult.ToString)
            If lowestResult.Delta > currentResult.Delta Then lowestResult.Become(currentResult)
        Next
        Console.WriteLine("Parsing of results completed . . .")
        Console.WriteLine("------------------------------------")
        Console.WriteLine("Result with lowest latency:")
        Console.WriteLine(lowestResult.ToString)
        Console.WriteLine("------------------------------------")
    End Sub

    Sub ToClipboard(ByVal value As String)
        Console.WriteLine("")
        OpenClipboard(IntPtr.Zero)
        Dim ptr As IntPtr = Marshal.StringToHGlobalUni(value)
        SetClipboardData(13, ptr)
        CloseClipboard()
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine("String has been copied to clipboard.")
    End Sub

    Function ReadFile() As String
        Dim temp As String
        Using sReader As New StreamReader("results.txt")
            temp = sReader.ReadToEnd
            sReader.Close()
        End Using
        Return temp
    End Function
End Module
