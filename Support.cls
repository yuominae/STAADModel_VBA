VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Support"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Public fields
Public ID As Long
Public Nodes As Dictionary
Public SupportRelease As Release
Public SupportType As STAADSUPPORTTYPE

'' Methods
Private Sub Class_Initialize()
    
    Me.SupportType = STAADSUPPORTTYPE.UNSPECIFIED
    Set Me.Nodes = New Dictionary
    
End Sub

Public Sub Initialise(ByVal ID As Long)

    Me.ID = ID
    
End Sub
