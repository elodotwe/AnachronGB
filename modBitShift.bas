Attribute VB_Name = "modBitShift"
Public Function LowByte(value As Long) As Byte
    LowByte = value And &HFF
End Function

Public Function HighByte(value As Long) As Byte
    HighByte = RShift(value, 8) And &HFF
End Function

Public Function SetLowByte(value As Long, newByte As Byte) As Long
    Dim result As Long
    result = value And &HFF00
    result = result + newByte
    SetLowByte = result
End Function

Public Function SetHighByte(value As Long, newByte As Byte) As Long
    Dim result As Long
    result = value And &HFF
    result = result + newByte * &H100&
    SetHighByte = result
End Function

Public Function LShift(value, n As Integer)
    LShift = value * 2& ^ n
End Function

Public Function RShift(value, n As Integer)
    ' \ instead of / means integer division
    RShift = value \ 2& ^ n
End Function

