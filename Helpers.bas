Attribute VB_Name = "Helpers"
Option Explicit

Public Function Asin(ByVal value As Double) As Double

    Asin = Math.Atn(value / Sqr(-value * value + 1))

End Function

Public Function Acos(ByVal value As Double) As Double

    Acos = WorksheetFunction.Degrees(WorksheetFunction.Acos(value))

End Function
