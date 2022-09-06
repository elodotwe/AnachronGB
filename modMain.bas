Attribute VB_Name = "modMain"
Option Explicit

Sub main()
    Dim decoder As CPUInstructionDecoder
    Set decoder = New CPUInstructionDecoder
    decoder.runTests

    Dim state As CPURegisterState
    Set state = New CPURegisterState
    state.runTests
End Sub
