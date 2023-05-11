Attribute VB_Name = "Truby"
'===============================================================================
'   Макрос          : Truby
'   Версия          : 2022.05.11
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "Truby"

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    Dim Page As Page
    With InitData.RequestDocumentOrPage
        If .IsError Then Exit Sub
        Set Page = .Page
    End With
    
    BoostStart APP_NAME, RELEASE
    
    With New Main
        .StartRoutine Page
    End With
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # тесты

Private Sub testSomething()
    Dim Element As Variant
    For Each Element In ActiveShape.Style.GetAllPropertyNames
        Show Element
    Next Element
End Sub

Private Sub testStyle()
    Dim Shape As Shape
    For Each Shape In ActivePage.Shapes
        If Shape.Type = cdrLinearDimensionShape Then
            Show Shape.Style.GetPropertyAsString("dimension")
        End If
    Next Shape
End Sub
