VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String

Public AssignedBeams As Dictionary

Private Sub Class_Initialize()

    Set Me.AssignedBeams = New Dictionary

End Sub
