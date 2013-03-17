Module MainModule

    'declare the timer to loop the procedure and set its interval to 2 seconds
    Dim WithEvents looper As New System.Timers.Timer(2000)
    Dim WithEvents looper2 As New System.Timers.Timer(30000)

    'declare the constants
    Const TotalOvers As Integer = &H930EB4      'Total overs in the match
    Const OversBowled As Integer = &H8175DC     'overs bowled in the current innings
    Const Runs As Integer = &HF30164            'runs in the current innings

    'declare the trainer object used to read/write memory and set its window name
    Dim trnr As New Trainer("EA SPORTS™ Cricket 07")

    Dim checkPoint As Byte                      'Next check point (overs) for new configs to load
    Dim part As Byte                            'the current part
    Dim partValues(3) As Single
    Dim TableCPU(14, 1) As Single               'contains ai values and offsets of the CPU
    Dim TableHuman(14, 1) As Single             '      "       "       "       "       Human
    Dim prevMod As Single                       'stores the previous runrate modification to subtract
    Dim RR As Single

    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


    Sub Main()

        'load the partitions from the parts file
        'Dim _f As New IO.FileStream("Parts.txt", IO.FileMode.Open, IO.FileAccess.Read)
        Dim _r As New IO.StreamReader("Parts.txt")
        For i As Byte = 0 To 3
            partValues(i) = _r.ReadLine
        Next
        _r.Close()
        '_f.Close()
        '_r.Dispose()
        '_f.Dispose()
        _r = New IO.StreamReader("fname.txt")

        _r.Close()
        _r.Dispose()

        Console.WriteLine("Gameplay07 Launched, Now Start Cricket 07 and Enjoy")
        Console.WriteLine(CStr(trnr.GetSValue(TotalOvers)))

        'populate the offsets of the arrays that hold the offsets and value table
        PopulateArrays()

        'start the loop
        looper.Start()
        looper2.Start()

        'keep the application alive
        Console.Read()

    End Sub

    Function RunRate(ByVal runs As Integer, ByVal overs As Single) As Single

        'returns the current runrate.
        'note that the program can't detect the number of balls bowled in the over
        'so on an average it is taken as 3 balls bowled in the over.
        'With trnr
        Return runs / (overs + 0.3)
        'End With

    End Function

    Private Sub looper_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles looper.Elapsed

        'check if control key is pressed,
        'if pressed a new inning is started,
        'so set the default settings
        If GetAsyncKeyState(17) <> 0 Then
            part = 0
            ModifyConfigs()
            UpdateConfigs()
            setCheckPoint()
        End If

        checkPointElapsed()

    End Sub

    Private Sub setCheckPoint()

        'set the next check point by multiplying the required value by the
        'total overs in an inning in the current match
        checkPoint = Math.Round(partValues(part) * trnr.GetSValue(TotalOvers), 0)

    End Sub

    Private Function checkPointElapsed() As Boolean

        'check if the overs bowled are equal to or more than the check point,
        'if true then increment the part and set a new check point and modify & update configs
        If trnr.GetSValue(OversBowled) >= checkPoint Then
            part += 1
            ModifyConfigs()
            UpdateConfigs()
            setCheckPoint()
            Return True
        End If
        Return False

    End Function

    Private Sub PopulateArrays()

        Dim buffer As String, counter As Byte
        Dim _r As New IO.StreamReader("Human.g07")
        Do While _r.EndOfStream <> True
            buffer = _r.ReadLine
            If Not Mid(buffer, 1, 1) = "-" Then
                counter += 1
                TableHuman(counter - 1, 0) = CSng(buffer)
            End If
        Loop

        counter = 0
        _r.Close()
        _r = New IO.StreamReader("CPU.g07")
        Do While _r.EndOfStream <> True
            buffer = _r.ReadLine
            If Not Mid(buffer, 1, 1) = "-" Then
                counter += 1
                TableCPU(counter - 1, 0) = CSng(buffer)
            End If
        Loop
        _r.Close()
        _r.Dispose()

    End Sub

    Private Sub UpdateConfigs()

        If trnr.OpenProcess(True) = True Then
            For i As Byte = 0 To 13
                trnr.SetValue(CInt(TableHuman(i, 0)), TableHuman(i, 1))
                trnr.SetValue(CInt(TableCPU(i, 0)), TableCPU(i, 1))
            Next
            trnr.CloseProcess()
            Console.WriteLine("Configs Updated")
        End If

    End Sub

    Private Sub ModifyConfigs()
        On Error Resume Next

        Dim totOvers As Byte = trnr.GetSValue(TotalOvers)

        Dim _r As New IO.StreamReader(".\Config\" & totOvers.ToString & "\HumanPart" & part.ToString & ".g07")
        For i As Byte = 0 To 14
            TableHuman(i, 1) = CSng(_r.ReadLine)
        Next
        _r.Close()

        _r = New IO.StreamReader(".\Config\" & totOvers.ToString & "\CPUPart" & part.ToString & ".g07")
        For i As Byte = 0 To 14
            TableCPU(i, 1) = CSng(_r.ReadLine)
        Next

    End Sub

    Private Sub ModRunRate()

        If trnr.OpenProcess(True) = True Then
            RR = RunRate(trnr.GetSValue(Runs), trnr.GetSValue(OversBowled))
            trnr.SetValue(TableHuman(0, 0), CSng(trnr.GetSValue(TableHuman(1, 0)) + RR * 0.4 - prevMod * 0.4))
            trnr.SetValue(TableHuman(3, 0), CSng(trnr.GetSValue(TableHuman(3, 0)) + RR * 0.2 - prevMod * 0.2))

            trnr.SetValue(TableCPU(0, 0), CSng(trnr.GetSValue(TableCPU(1, 0)) + RR * 0.4 - prevMod * 0.4))
            trnr.SetValue(TableCPU(3, 0), CSng(trnr.GetSValue(TableCPU(3, 0)) + RR * 0.15 - prevMod * 0.15))

            prevMod = RR

            trnr.CloseProcess()
            Console.WriteLine("Modded RunRate")
        End If

    End Sub

    Private Sub sBatsmanMode()

        Randomize()
        Dim i As Byte
        i = Rnd() * 4
        If i = 3 Then
            If trnr.OpenProcess(True) = True Then
                trnr.SetValue(TableHuman(3, 0), CSng(trnr.GetSValue(TableHuman(3, 0)) + 1.5))
                trnr.SetValue(TableCPU(3, 0), CSng(trnr.GetSValue(TableCPU(3, 0)) + 1.5))
                Console.WriteLine("sBatsman Mode Enabled")
                Exit Sub
            End If
        End If
        Console.WriteLine("sBatsman Mode Disabled")
    End Sub

    Private Sub looper2_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles looper2.Elapsed

        ModRunRate()
        sBatsmanMode()

    End Sub
End Module
