VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPUInstructionDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function decodeOpcode(address As Long, data() As Byte) As CPUInstruction
    Dim result As CPUInstruction
    Set result = decodeLDR8R8(address, data)
    If result Is Nothing Then
        Set result = decodeLDR8Immediate(address, data)
    End If
    Set decodeOpcode = result
End Function

Public Function runTests()
    testDecodeLDR8R8
    testDecodeLDR8Immediate
End Function

Private Function testDecodeLDR8R8()
    Dim test(0 To 2) As Byte
    'ld b, a
    test(0) = &H47
    'cruft
    test(1) = 2
    'cruft
    test(2) = 3
    Dim result As CPUInstruction
    Set result = decodeOpcode(0, test)
    Dim resultR8R8 As CPUInstruction_LD_R8_R8
    Set resultR8R8 = result
    AssertEquals CPURegister8_a, resultR8R8.sourceReg
    AssertEquals CPURegister8_b, resultR8R8.destReg
    
    'ld b, invalid (should decode to nothing)
    test(0) = &H46
    Set result = decodeOpcode(0, test)
    Assert result Is Nothing
End Function

Private Function decodeLDR8R8(address As Long, data() As Byte) As CPUInstruction
    Dim result As CPUInstruction_LD_R8_R8
    
    Dim sourceReg As CPURegister8
    Dim destReg As CPURegister8
    
    destReg = register8FromOpcodeField(RShift(data(address), 3) And &H7)
    sourceReg = register8FromOpcodeField(data(address) And &H7)
    
    If ((data(address) And &HC0) = &H40) And (Not (destReg = CPURegister8_invalid)) And (Not (sourceReg = CPURegister8_invalid)) Then
        Set result = New CPUInstruction_LD_R8_R8
        result.sourceReg = sourceReg
        result.destReg = destReg
    End If
    Set decodeLDR8R8 = result
End Function

Private Function testDecodeLDR8Immediate()
    Dim test(0 To 2) As Byte
    'ld b, n
    test(0) = &H6
    'n = 2
    test(1) = 2
    'cruft
    test(2) = 3
    Dim result As CPUInstruction_LD_R8_Immediate
    Set result = decodeOpcode(0, test)
    AssertEquals 2, result.sourceValue
    AssertEquals CPURegister8_b, result.destReg
End Function

Private Function decodeLDR8Immediate(address As Long, data() As Byte) As CPUInstruction
    Dim result As CPUInstruction_LD_R8_Immediate
    Dim sourceValue As Byte
    Dim destReg As CPURegister8
    sourceValue = data(address + 1)
    destReg = register8FromOpcodeField(RShift(data(address), 3) And &H7)
    
    If ((data(address) And &HC0) = &H40) And (Not (destReg = CPURegister8_invalid)) Then
        Set result = New CPUInstruction_LD_R8_Immediate
        result.sourceValue = sourceValue
        result.destReg = destReg
    End If
    Set decodeLDR8Immediate = result
End Function
