' Cross-platform VBA implementation of a high-precision timer.
' (Works on Windows and on macOS)
'
' Author: Guido Witt-Dörring
' Created: 2023/04/03
' Updated: 2023/05/16
' License: MIT
'
' ————————————————————————————————————————————————————————————————
' https://gist.github.com/guwidoe/5c74c64d79c0e1cd1be458b0632b279a
' ————————————————————————————————————————————————————————————————
'
' Copyright (c) 2023 Guido Witt-Dörring
'
' MIT License:
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.

Option Explicit

#If Mac Then
    #If VBA7 Then
        'https://developer.apple.com/documentation/kernel/1462446-mach_absolute_time
        Private Declare PtrSafe Function mach_continuous_time Lib "/usr/lib/libSystem.dylib" () As Currency
        Private Declare PtrSafe Function mach_timebase_info Lib "/usr/lib/libSystem.dylib" (ByRef timebaseInfo As MachTimebaseInfo) As Long
    #Else
        Private Declare Function mach_continuous_time Lib "/usr/lib/libSystem.dylib" () As Currency
        Private Declare Function mach_timebase_info Lib "/usr/lib/libSystem.dylib" (ByRef timebaseInfo As MachTimebaseInfo) As Long
    #End If
#Else
    #If VBA7 Then
        'https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter
        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As LongPtr
        Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As LongPtr
    #Else
        Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef Frequency As Currency) As Long
        Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef Counter As Currency) As Long
    #End If
#End If

#If Mac Then
    Private Type MachTimebaseInfo
        Numerator As Long
        Denominator As Long
    End Type
#End If

Public Enum TimeUnit
    tuSeconds = 0
    tuMilliseconds
    tuMicroseconds
    tuAutomatic
End Enum

'Returns operating system clock tick count since system startup
Private Function GetTickCount() As Currency
    #If Mac Then
        GetTickCount = mach_continuous_time()
    #Else
        QueryPerformanceCounter GetTickCount
    #End If
End Function

'Returns frequency in ticks per second
Private Function GetFrequency() As Currency
    Static isInitialized As Boolean
    Static freqency As Currency
    If Not isInitialized Then
        #If Mac Then
            Dim tbInfo As MachTimebaseInfo: mach_timebase_info tbInfo
    
            freqency = (tbInfo.Denominator / tbInfo.Numerator) * 100000@
        #Else
            QueryPerformanceFrequency freqency
        #End If
        isInitialized = True
    End If
    GetFrequency = freqency
End Function

'Returns time since system startup in seconds with 0.1ms (=100µs) precision
Public Function AccurateTimer(Optional ByVal unit As TimeUnit = tuSeconds) _
                                       As Currency
    Select Case unit
        Case tuMicroseconds: AccurateTimer = AccurateTimerUs
        Case tuMilliseconds, tuAutomatic: AccurateTimer = AccurateTimerMs
        Case tuSeconds: AccurateTimer = AccurateTimerS
    End Select
End Function

'Returns time since system startup in seconds with 0.1ms (=100µs) precision
Public Function AccurateTimerS() As Currency
    AccurateTimerS = GetTickCount / GetFrequency
End Function

'Returns time since system startup in milliseconds with 0.1µs (=100ns) precision
Public Function AccurateTimerMs() As Currency
    'Note that this calculation will work even if 1000@ / GetFrequency < 0.0001
    AccurateTimerMs = (1000@ / GetFrequency) * GetTickCount
End Function

'Returns time since system startup in microseconds, up to 0.1ns =100ps precision
'The highest precision achieved by this function depends on the system, however,
'typically precision will be the same as for AccurateTimerMs.
Public Function AccurateTimerUs() As Currency
    AccurateTimerUs = (1000000@ / GetFrequency) * GetTickCount
End Function

'Starts/resets a timer in the background
Public Sub StartTimer(Optional ByVal printHeaders As Boolean = True)
    TimerBackend 0, printHeaders
End Sub
Public Sub st(Optional ByVal printHeaders As Boolean = True)
    TimerBackend 0, printHeaders
End Sub

'Resets/starts the timer in the background, alias for StartTimer
Public Sub ResetTimer(Optional ByVal printHeaders As Boolean = False)
    TimerBackend 0, printHeaders
End Sub

'Seconds to Microseconds conversion function for convenient checking against
'return values of 'ReadTimer'
Public Function StoUs(ByRef s As Currency) As Currency
    StoUs = s * 1000000
End Function

'Prints the time that has passed since the last `StartTimer` or `ResetTimer`
'has been called to the immediate window and the `description` next to it.
'This sub by default subtracts its own runtime from the current timers total
'time to avoid skewing the timing results of profiled code. If that is not
'desired, i.e. for other applications than code profiling, call it with
'`subtractOwnRuntime = False`
'If 'unit = tuAutomatic', the return value is always in µs
Public Function ReadTimer(Optional ByRef description As String = vbNullString, _
                          Optional ByVal unit As TimeUnit = tuAutomatic, _
                          Optional ByVal reset As Boolean = False, _
                          Optional ByVal subtractOwnRuntime As Boolean = True, _
                          Optional ByVal printResult As Boolean = True) _
                                   As Currency
    ReadTimer = TimerBackend(1, description, unit, reset, subtractOwnRuntime, _
                             printResult)
End Function
Public Function RT(Optional ByRef description As String = vbNullString, _
                   Optional ByVal unit As TimeUnit = tuAutomatic, _
                   Optional ByVal reset As Boolean = False, _
                   Optional ByVal subtractOwnRuntime As Boolean = True, _
                   Optional ByVal printResult As Boolean = True) _
                            As Currency
    RT = TimerBackend(1, description, unit, reset, subtractOwnRuntime, _
                      printResult)
End Function

Private Function TimerBackend(ByVal command As Long, _
                              ParamArray arr() As Variant) As Currency
    Static timeStamp As Currency
    Static callsSinceReset As Long
    
    'Always do this first for maximum accuracy
    Dim timeAtCall As Currency: timeAtCall = GetTickCount
    Dim readTimeUs As Currency: readTimeUs = (1000000@ / GetFrequency) * _
                                             (timeAtCall - timeStamp)
    Select Case command
        Case 0 'StartTimer or ResetTimer
            If arr(0) Then 'if printHeaders ...
                Debug.Print "Time taken", "Task description"
            End If
            callsSinceReset = 0
            timeStamp = GetTickCount
        Case 1 'ReadTimer
            callsSinceReset = callsSinceReset + 1
            Dim description As String: description = arr(0)
            If description = "" Then description = "Task " & callsSinceReset
            Dim unit As TimeUnit: unit = arr(1)
            If unit = tuAutomatic Then
                Select Case readTimeUs
                    Case Is > 1000000: unit = tuSeconds
                    Case 1000 To 1000000: unit = tuMilliseconds
                    Case Else: unit = tuMicroseconds
                End Select
            End If
            Select Case unit 'Unit
                Case TimeUnit.tuSeconds: TimerBackend = readTimeUs / 1000000@
                Case TimeUnit.tuMilliseconds: TimerBackend = readTimeUs / 1000@
                Case TimeUnit.tuMicroseconds: TimerBackend = readTimeUs
            End Select
            If arr(4) Then 'If printResult ...
                Debug.Print TimerBackend & Choose(unit + 1, " s", " ms", _
                            IIf("µ" = Chr$(181), " µs", " us")), description
            End If
            'If unit was tuAutomatic, override return unit to ensure consistency
            If arr(1) Then TimerBackend = readTimeUs
            If arr(2) Then 'If reset ...
                callsSinceReset = 0
                timeStamp = GetTickCount
            Else
                'Subtract runtime of this method from future `ReadTimer` calls
                If arr(3) Then _
                    timeStamp = timeStamp + (GetTickCount - timeAtCall)
            End If
    End Select
End Function

'———————————————————————————————————————————————————————————————————————————————
'                                  DEMO PART
'———————————————————————————————————————————————————————————————————————————————
'This demonstrates the simplest, and the recommended way the procedures provided
'in this module can be used to profile your code:
Sub DemoCodeExecutionTiming()
    StartTimer
    
    'Some code that does something, e.g.:
    Dim i As Long
    For i = 1 To 100000
    Next i
    
    ReadTimer "Looping 100000 times." 'The desctiption is optional
End Sub

'This is a way of using the provided `AccurateTimer` functions to time things
'in a way that mimics how the built in `Timer` function is commonly used
Private Sub DemoAccurateTimer()
    Dim s As Currency:  s = AccurateTimerS   'or: AccurateTimer(tuSeconds)
    Dim ms As Currency: ms = AccurateTimerMs 'or: AccurateTimer(tuMilliseconds)
    Dim µs As Currency: µs = AccurateTimerUs 'or: AccurateTimer(tuMicroseconds)

    Dim i As Long
    For i = 1 To 10000000
        i = i
    Next i
    
    Debug.Print "Code execution took " & AccurateTimerS - s & " seconds."
    Debug.Print "Code execution plus time of the first 'Debug.Print' statement: " _
                & AccurateTimerMs - ms & " milliseconds."
    Debug.Print "Code execution plus time of the first two 'Debug.Print' " & _
                "statements: " & AccurateTimerUs - µs & " microseconds."
End Sub

'———————————————————————————————————————————————————————————————————————————————
'     SYSTEM SPECIFIC PERFORMANCE TESTING AND EXPLANATORY DEMONSTRATIONS
'———————————————————————————————————————————————————————————————————————————————
'This sub runs all the following subs, the results are printed to the immediate
'window
Sub RunAll()
    ShowAverageDelayInTimingCausedByReadTimerCall
    DemoSubtractOwnRuntime
    ShowPrecisionOfTimersOnCurrentSystem
End Sub

'Even though `ReadTimer` by default subtracts its own runtime from the total
'time, a tiny overhead caused by the (API) function calls themselves does occur.
'This Sub demonstrated the various delays in the timing data on your own system,
'depending on how `ReadTimer` is called.
Sub ShowAverageDelayInTimingCausedByReadTimerCall()
    Const LOOP_COUNT As Long = 100000
    Dim i As Long
    
    StartTimer
    For i = 1 To LOOP_COUNT / 1000
        ReadTimer printResult:=True, subtractOwnRuntime:=False
    Next i
    Debug.Print "Average delay in timing data caused by ReadTimer calls: " & _
     vbNewLine & "If `subtractOwnRuntime = False` and `printResult:=True`: " & _
           ReadTimer(printResult:=False) / (LOOP_COUNT / 1000) & " microseconds"
           
    ResetTimer
    For i = 1 To LOOP_COUNT / 10
        ReadTimer printResult:=False, subtractOwnRuntime:=False
    Next i
    Debug.Print "If `subtractOwnRuntime = False` and `printResult:=False`: " & _
             ReadTimer(printResult:=False) / (LOOP_COUNT / 10) & " microseconds"
             
    ResetTimer
    For i = 1 To LOOP_COUNT
        ReadTimer printResult:=False
    Next i
    Debug.Print "If `subtractOwnRuntime = True`, regardless of `printResult`: " _
                & ReadTimer(printResult:=False) / LOOP_COUNT & " microseconds"

End Sub

'This Sub demonstrates why the `subtractOwnRuntime` feature is useful.
Private Sub DemoSubtractOwnRuntime()
    Const LOOP_COUNT As Long = 100000
    Dim loopTimeTaken As Currency
    
    Debug.Print vbNewLine & "Some simple code timing:" & vbNewLine
    StartTimer
    
    Dim i As Long
    For i = 1 To LOOP_COUNT
    Next i
    
    loopTimeTaken = ReadTimer(printResult:=False)
    ReadTimer "Time to loop " & LOOP_COUNT & " times."
    ReadTimer "in seconds", tuSeconds
    ReadTimer "in milliseconds", tuMilliseconds
    ReadTimer "in microseconds", tuMicroseconds
    
    Debug.Print ""
    Debug.Print "The reason the time values printed by `ReadTimer` are much" _
    & vbNewLine & "closer together than the time a `Debug.Print` statement" & _
    vbNewLine & "usually takes is, that `ReadTimer` subtracts its own runtime" _
    & vbNewLine & "from the currently running timers total time."
    Debug.Print "Perhaps the following illustrates why that is beneficial:"
    
    Dim debugPrintTimeTaken As Currency
    Debug.Print vbCrLf; "Test Debug.Print statements to see how long it takes:"
    
    ResetTimer
    Debug.Print "TEST"
    debugPrintTimeTaken = ReadTimer("... time taken to print ""TEST""")
    
    Debug.Print vbNewLine & "The single `Debug.Print` statement took about " & _
                CCur(debugPrintTimeTaken / loopTimeTaken) & " times " & _
                vbNewLine & "as long as looping " & LOOP_COUNT & " times." & vbLf
End Sub

'Prints the precisions of the above timer functions for the current system
Private Sub ShowPrecisionOfTimersOnCurrentSystem()
    Dim dtp As Currency: dtp = 0.1 'DataType precision (in ms/µs/ns)
    Dim F As Currency:   F = GetFrequency 'Frequency
    Debug.Print "Precision for timing unit 'tuSeconds' is " & _
                IIf(dtp > (0.1@ / F), dtp, (0.1@ / F)) & " milliseconds."
    Debug.Print "Precision for timing unit 'tuMilliseconds' is " & _
                IIf(dtp > (100@ / F), dtp, (100@ / F)) & " microseconds."
    Debug.Print "Precision for timing unit 'tuMicroseconds' is " & _
                IIf(dtp > (100000@ / F), dtp, (100000@ / F)) & " nanoseconds."
End Sub
