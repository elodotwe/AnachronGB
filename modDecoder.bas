Attribute VB_Name = "modDecoder"
Option Explicit

Enum CPURegister8
    a
    b
    c
    d
    e
    h
    l
End Enum

Enum CPURegister16
    af
    bc
    de
    hl
    sp
    pc
End Enum

Public Function register8Name(register As CPURegister8)
    Dim r As String
    
    Select Case register
    Case a: r = "a"
    Case b: r = "b"
    Case c: r = "c"
    End Select
    register8Name = r
End Function


Public Function decodeOpcode(address As Long, data() As Byte) As CPUInstruction
    Dim result As CPUInstruction
    Set result = decodeLDR8R8(address, data)
    If result Is Nothing Then
        MsgBox ("not an ldr8r8")
    End If
    
    Set decodeOpcode = result
End Function


Private Function decodeLDR8R8(address As Long, data() As Byte) As CPUInstruction
    Dim result As CPUInstruction_LD_R8_R8
    If data(address) = 1 Then
        Set result = New CPUInstruction_LD_R8_R8
        result.sourceReg = a
        result.destReg = c
    End If
    Set decodeLDR8R8 = result
End Function
