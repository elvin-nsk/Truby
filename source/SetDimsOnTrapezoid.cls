VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetDimsOnTrapezoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Declarations

Private Type typeThis
    LineColor As Color
    LineWidth As Double
    TextColor As Color
    TextSize As Double
End Type
Private This As typeThis

'===============================================================================
' # Events

Public Event OnRectangleCreation(ByVal Rect As Rect)

'===============================================================================
' # Handlers

Private Sub Class_Initialize()
    With This
        Set .LineColor = CreateCMYKColor(0, 0, 0, 100)
        .LineWidth = 1
        Set .TextColor = CreateCMYKColor(0, 0, 0, 100)
        .TextSize = 100
    End With
End Sub

Private Sub Class_Terminate()
'
End Sub

'===============================================================================
' # Public

Public Sub SetSizes( _
               ByVal LineWidth As Double, _
               ByVal TextSize As Double _
           )
    This.LineWidth = LineWidth
    This.TextSize = TextSize

End Sub

Public Function MakeDims( _
                    ByVal Trapezoid As Shape _
                ) As Shape
    Dim Rectangle As Shape
    Set Rectangle = MakeRectangleAroundTrapezoid(Trapezoid)
    Rectangle.Outline.SetNoOutline
    With Rectangle.SnapPointsOfType(cdrSnapPointEdge)
        MakeDim _
            Rectangle, .Item(1), .Item(2)
        MakeDim _
            Rectangle, .Item(2), .Item(3)
    End With
End Function

'===============================================================================
' # Helpers

Private Function MakeDim( _
                     ByVal Shape As Shape, _
                     ByVal Point1 As SnapPoint, _
                     ByVal Point2 As SnapPoint _
                 ) As Shape
    Set MakeDim = _
        ActiveLayer.CreateLinearDimension( _
            Type:=cdrDimensionSlanted, _
            Point1:=Point1, _
            Point2:=Point2, _
            TextCentered:=True, _
            TextX:=Shape.LeftX - 80, _
            TextY:=Shape.TopY + 50, _
            Style:=cdrDimensionStyleDecimal, _
            Precision:=0, _
            ShowUnits:=False, _
            Units:=cdrDimensionUnitMM, _
            Placement:=cdrDimensionAboveLine, _
            HorizontalText:=False, _
            BoxedText:=False, _
            LeadingZero:=True, _
            OutlineWidth:=This.LineWidth, _
            OutlineColor:=This.LineColor, _
            TextSize:=This.TextSize, _
            TextColor:=This.TextColor _
        )
    'при построении часть свойств корелом игнорится
    'и берётся из стиля в документе
    'вот так можно оверрайднуть
    With MakeDim.Style.GetProperty("dimension")
        .SetProperty "precision", 0
        .SetProperty "showUnits", 0
        .SetProperty "textPlacement", 0
        .SetProperty "units", 3
    End With
End Function

Private Function MakeRectangleAroundTrapezoid( _
                     ByVal Trapezoid As Shape _
                 ) As Shape
    Dim LongestSeg As Segment
    Set LongestSeg = FindLongestSegment(Trapezoid.DisplayCurve)
    Dim Angle As Double
    Angle = LongestSeg.EndingControlPointAngle
    Trapezoid.RotateEx -Angle, 0, 0
    Dim Result As Shape
    Set Result = ActiveLayer.CreateRectangleRect(Trapezoid.BoundingBox)
    RaiseEvent OnRectangleCreation(Result.BoundingBox.GetCopy)
    Trapezoid.RotateEx Angle, 0, 0
    Result.RotateEx Angle, 0, 0
    Set MakeRectangleAroundTrapezoid = Result
End Function

Private Function FindLongestSegment( _
                     ByVal Curve As Curve _
                 ) As Segment
    Dim Seg As Segment
    Set FindLongestSegment = Curve.Segments.First
    For Each Seg In Curve.Segments
        If Seg.Length > FindLongestSegment.Length Then
            Set FindLongestSegment = Seg
        End If
    Next Seg
End Function

Private Function xxx( _
                     ByVal a As IPoint, _
                     ByVal b As IPoint, _
                     ByVal c As IPoint, _
                     ByVal D As IPoint _
                 ) As Shape

End Function
