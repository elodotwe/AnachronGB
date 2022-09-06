VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPUInstruction_LD_R8_Immediate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CPUInstruction

Public sourceValue As Byte
Public destReg As CPURegister8

Public Property Get CPUInstruction_DebugDescription() As String
    CPUInstruction_DebugDescription = "LD " + register8Name(destReg) + "," + sourceValue
End Property

Public Property Get CPUInstruction_OpcodeLength() As Integer
    CPUInstruction_OpcodeLength = 2
End Property

