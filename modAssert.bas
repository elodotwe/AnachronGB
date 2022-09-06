Attribute VB_Name = "modAssert"
Public Sub AssertEquals(expected, actual)
    Assert expected = actual, "Expected " + Hex(expected) + " == " + Hex(actual)
End Sub

Public Sub Assert(condition As Boolean, Optional description As String)
    If Not condition Then
        Err.Raise vbObjectError + 1000, "Unit tests", "Assertion failed: " + description
    End If
End Sub
