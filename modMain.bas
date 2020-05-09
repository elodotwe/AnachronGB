Attribute VB_Name = "modMain"
Option Explicit

Dim RunTests As Boolean

Sub main()
    RunTests = True

    If RunTests Then
        Dim test(0 To 2) As Byte
        test(0) = 1
        test(1) = 2
        test(2) = 3
        Dim result As CPUInstruction
        Set result = modDecoder.decodeOpcode(0, test)
        If result Is Nothing Then
            MsgBox ("null result")
        Else
            MsgBox ("instruction was " + result.DebugDescription)
        End If
        
    
        Dim state As CPURegisterState
        Set state = New CPURegisterState
        state.RunTests
    End If
    
End Sub
