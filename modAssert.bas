Attribute VB_Name = "modAssert"
Public Sub AssertEquals(expected As Long, actual As Long)
    Assert expected = actual, "Expected " + Hex(expected) + " == " + Hex(actual)
End Sub

Public Sub Assert(condition As Boolean, Optional description As String)
    If Not condition Then
        Err.Raise vbObjectError + 1000, "CPURegisterState unit tests", "Assertion failed: " + description
    End If
End Sub
