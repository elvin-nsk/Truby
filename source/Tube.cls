VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Struct
Option Explicit

'===============================================================================

Public Length As Long
Public Diameter As Long
Public Quantity As Long

'===============================================================================

Public Property Get Self() As Tube
    Set Self = Me
End Property
