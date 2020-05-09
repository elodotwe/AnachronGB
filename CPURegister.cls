VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPURegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public value As Long


Public Property Get high() As Byte
    high = HighByte(value)
End Property

Public Property Let high(ByVal newVal As Byte)
    value = SetHighByte(high, newVal)
End Property

Public Property Get low() As Byte
    low = LowByte(value)
End Property

Public Property Let low(ByVal newVal As Byte)
    value = SetLowByte(value, newVal)
End Property
