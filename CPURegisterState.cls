VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPURegisterState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public af As CPURegister
Public bc As CPURegister
Public de As CPURegister
Public hl As CPURegister
Public sp As CPURegister
Public pc As CPURegister

Public Property Get FLAGS_ZF()
    FLAGS_ZF = &H80
End Property

Public Property Get FLAGS_N()
    FLAGS_N = &H40
End Property

Public Property Get FLAGS_H()
    FLAGS_H = &H20
End Property

Public Property Get FLAGS_CY()
    FLAGS_CY = &H10
End Property


Public Sub RunTests()
    AssertEquals 8, LShift(1, 3)
    AssertEquals 1, RShift(10, 3)
    
    af.value = &HFFFF&
    af.high = &HEE
    AssertEquals af.value, &HEEFF&
    AssertEquals af.high, &HEE
    AssertEquals af.low, &HFF
    
    af.low = &H11
    AssertEquals af.value, &HEE11&
    AssertEquals af.high, &HEE
    AssertEquals af.low, &H11
End Sub

Private Sub Class_Initialize()
    Set af = New CPURegister
    Set bc = New CPURegister
    Set de = New CPURegister
    Set hl = New CPURegister
    Set sp = New CPURegister
    Set pc = New CPURegister
End Sub
