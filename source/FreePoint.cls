VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FreePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@ModuleDescription "Low-coupled version from Core framework"
'@PredeclaredId
Option Explicit

Implements IPoint

'===============================================================================

Private Type typeThis
  x As Double
  Y As Double
  Frozen As Boolean
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal x As Double, ByVal Y As Double) As IPoint
  With New FreePoint
    Set Create = .Init(x, Y)
  End With
End Function

Friend Function Init(ByVal x As Double, ByVal Y As Double) As IPoint
  With This
    .x = x
    .Y = Y
  End With
  Set Init = Me
End Function

'===============================================================================

Private Function IPoint_GetCopy() As IPoint
  Set IPoint_GetCopy = Create(This.x, This.Y)
End Function

Private Function IPoint_GetDistanceFrom(ByVal Point As IPoint) As Double
End Function

Private Sub IPoint_Move(ByVal DeltaX As Double, ByVal DeltaY As Double)
  With This
    .x = .x + DeltaX
    .Y = .Y + DeltaY
  End With
End Sub

Private Sub IPoint_RotateAroundPoint(ByVal RotationCenter As IPoint, ByVal Angle As Double)
  With RotationCenter
    This.x = (This.x - .x) * VBA.Cos(DegToRad(Angle)) - _
             (This.Y - .Y) * VBA.Sin(DegToRad(Angle)) + .x
    This.Y = (This.x - .x) * VBA.Sin(DegToRad(Angle)) + _
             (This.Y - .Y) * VBA.Cos(DegToRad(Angle)) + .Y
  End With
End Sub

Private Property Get IPoint_Self() As IPoint
  Set IPoint_Self = Me
End Property

Private Property Get IPoint_X() As Double
  IPoint_X = This.x
End Property
Private Property Let IPoint_X(ByVal RHS As Double)
  This.x = RHS
End Property

Private Property Get IPoint_Y() As Double
  IPoint_Y = This.Y
End Property
Private Property Let IPoint_Y(ByVal RHS As Double)
  This.Y = RHS
End Property

'===============================================================================

Private Function DegToRad(Degree As Double) As Double
  DegToRad = Degree * 3.14159265358979 / 180
End Function

Private Property Get Self() As IPoint
  Set Self = Me
End Property
