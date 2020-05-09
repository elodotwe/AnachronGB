VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPUInstruction_LD_R8_R8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CPUInstruction

Public sourceReg As CPURegister8
Public destReg As CPURegister8

Public Property Get CPUInstruction_DebugDescription() As String
    CPUInstruction_DebugDescription = "LD " + register8Name(sourceReg) + "," + register8Name(destReg)
End Property


