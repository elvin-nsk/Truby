VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Declarations

Private Type typeThis
    Factor As Double
    Page As Page
    TextSize As Double
    TubeRects As Collection
End Type
Private This As typeThis

Private WithEvents SetDimsRoutine As SetDimsOnTrapezoid
Attribute SetDimsRoutine.VB_VarHelpID = -1

'===============================================================================
' # Handlers

Private Sub Class_Initialize()
    Set SetDimsRoutine = New SetDimsOnTrapezoid
    With This
        Set .TubeRects = New Collection
    End With
End Sub

Private Sub Class_Terminate()
'
End Sub

Private Sub SetDimsRoutine_OnRectangleCreation(ByVal Rect As Rect)
    This.TubeRects.Add Rect
End Sub

'===============================================================================
' # Public

Public Sub StartRoutine(ByVal Page As Page)
    With This
        Set .Page = Page
        CalcSizes
        SetDims
        SetText
    End With
End Sub

Public Property Get Self() As Main
    Set Self = Me
End Property

'===============================================================================
' # Logic

Private Sub CalcSizes()
    With This
        .Factor = LesserSize(.Page.Shapes.All.BoundingBox)
        .TextSize = .Factor * 0.4
        SetDimsRoutine.SetSizes .Factor * 0.002, .TextSize, .Factor * 0.002
    End With
End Sub

Private Sub SetDims()
    Dim Shape As Shape
    For Each Shape In This.Page.Shapes.FindShapes
        If ShapeValidForDims(Shape) Then SetDimsRoutine.MakeDims Shape
    Next Shape
End Sub

Private Sub SetText()
    Dim Tubes As Collection
    Set Tubes = RectsToTubes
    'Set Tubes = SortTubes(Tubes)
    Dim Text As String
    Dim Tube As Tube
    For Each Tube In Tubes
        Text = Text _
             & Tube.Length & " " & VBA.ChrW(8212) & " " _
             & Tube.Quantity & " רע." & VBA.Space(5) _
             & Tube.Diameter & " " & VBA.ChrW(215) & vbCr
    Next Tube
    MakeText Text
End Sub

'===============================================================================
' # Helpers

Private Property Get ShapeValidForDims(ByVal Shape As Shape) As Boolean
    If Not Shape.Fill.Type = cdrUniformFill Then Exit Property
    If GetColorLightness(Shape.Fill.UniformColor) > 250 Then Exit Property
    If Not HasDisplayCurve(Shape) Then Exit Property
    ShapeValidForDims = True
End Property

Private Property Get RectsToTubes() As Collection
    Set RectsToTubes = New Collection
    Dim TubeRect As Rect
    For Each TubeRect In This.TubeRects
        AddTube RectsToTubes, TubeRect
    Next TubeRect
End Property

Private Sub AddTube(ByVal ioTubes As Collection, ByVal TubeRect As Rect)
    Dim Tube As Tube
    Set Tube = _
        FindOrCreateTube( _
            ioTubes, _
            VBA.CLng(MinOfTwo(TubeRect.Width, TubeRect.Height)), _
            VBA.CLng(MaxOfTwo(TubeRect.Width, TubeRect.Height)))
    Tube.Quantity = Tube.Quantity + 1
End Sub

Private Function FindOrCreateTube( _
                     ByVal ioTubes As Collection, _
                     ByVal Diameter As Long, _
                     ByVal Length As Long _
                 ) As Tube
    Dim Tube As Tube
    For Each Tube In ioTubes
        If Tube.Diameter = Diameter _
       And Tube.Length = Length Then
            Set FindOrCreateTube = Tube
            Exit Function
        End If
    Next Tube
    Set FindOrCreateTube = New Tube
    FindOrCreateTube.Diameter = Diameter
    FindOrCreateTube.Length = Length
    ioTubes.Add FindOrCreateTube
End Function

'todo
Private Property Get SortTubes( _
                         ByVal Tubes As Collection _
                     ) As Collection
    Set SortTubes = New Collection
    Dim Tube As Tube
    Dim LastLength As Long
    For Each Tube In Tubes
        If Tube.Length <= LastLength Then SortTubes.Add Tube
    Next Tube
End Property

Private Sub MakeText(ByVal Text As String)
    Dim Space As Double
    Space = This.Factor * 0.1
    Dim PageShapes As ShapeRange
    Set PageShapes = This.Page.Shapes.All
    With ActiveLayer.CreateArtisticText(0, 0, Text, Size:=This.TextSize)
        .LeftX = PageShapes.LeftX
        .TopY = PageShapes.BottomY - Space
        .Text.Story.LineSpacing = 150
    End With
End Sub
