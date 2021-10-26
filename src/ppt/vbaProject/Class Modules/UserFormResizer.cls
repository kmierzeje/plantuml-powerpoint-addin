VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserFormResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------------------------------
' Author    : Manfred van den Noort
' Copyright : © 2020 worksheetsvba.com, all rights reserved
' Date      : 2020-12-11
' Version   : 1.0
' Purpose   : Create Resizable UserForm
'--------------------------------------------------------------------------------------------------------------------------

Private WithEvents frmResizableForm As MSForms.UserForm
Attribute frmResizableForm.VB_VarHelpID = -1
Private oResizableForm As Object
Private WithEvents frResizerCorner As MSForms.Frame
Attribute frResizerCorner.VB_VarHelpID = -1
Private WithEvents frResizerRight As MSForms.Frame
Attribute frResizerRight.VB_VarHelpID = -1
Private WithEvents frResizerBottom As MSForms.Frame
Attribute frResizerBottom.VB_VarHelpID = -1
Private sngMinHeight As Single
Private sngMinWidth  As Single
Private sngMouseX As Single
Private sngMouseY As Single

Event Resizing(ByVal X As Single, ByVal Y As Single)

Friend Property Set ResizableForm(ByRef oFrm As Object)
    Set frmResizableForm = oFrm
    Set oResizableForm = oFrm
    'set default values for MinHeight and MinWidth if they are not already set or if the values are greater than the initial dimensions
    If sngMinHeight = 0 Or sngMinHeight > oResizableForm.height Then
        sngMinHeight = oResizableForm.height
    End If
    If sngMinWidth = 0 Or sngMinWidth > oResizableForm.width Then
        sngMinWidth = oResizableForm.width
    End If
    AddResizeControls
End Property

Friend Property Let MinHeight(sngValue As Single)
    If oResizableForm Is Nothing Then
        sngMinHeight = sngValue
    ElseIf sngValue = 0 Or sngValue > oResizableForm.height Then
        sngMinHeight = oResizableForm.height
    Else
        sngMinHeight = sngValue
    End If
End Property

Friend Property Let MinWidth(sngValue As Single)
    If oResizableForm Is Nothing Then
        sngMinWidth = sngValue
    ElseIf sngValue = 0 Or sngValue > oResizableForm.width Then
        sngMinWidth = oResizableForm.width
    Else
        sngMinWidth = sngValue
    End If
End Property

Private Sub AddResizeControls()
    'frames are used to asure the resize controls are always on top of other form controls
    Set frResizerCorner = oResizableForm.Controls.Add("Forms.Frame.1")
    With frResizerCorner
        .SpecialEffect = fmSpecialEffectFlat
        .MousePointer = fmMousePointerSizeNWSE
        .ZOrder 0
        .width = 15
        .height = 15
    End With
    With frResizerCorner.Add("Forms.label.1")
        With .Font
            .Name = "Marlett"
            .Charset = 2
            .Size = 14
            .Bold = True
        End With
        .Caption = "o"
        .ForeColor = 6579300
        .width = 14
        .height = 14
        .Top = 1
        .Left = 1
        .Enabled = False
    End With
    Set frResizerRight = oResizableForm.Controls.Add("Forms.Frame.1")
    With frResizerRight
        .SpecialEffect = fmSpecialEffectFlat
        .MousePointer = fmMousePointerSizeWE
        .ZOrder 0
        .width = 2
        .Top = 0
    End With
    Set frResizerBottom = oResizableForm.Controls.Add("Forms.Frame.1")
    With frResizerBottom
        .SpecialEffect = fmSpecialEffectFlat
        .MousePointer = fmMousePointerSizeNS
        .ZOrder 0
        .height = 2
        .Left = 0
    End With
End Sub

Private Sub frResizerCorner_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        sngMouseX = X
        sngMouseY = Y
    End If
End Sub

Private Sub frResizerCorner_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        With oResizableForm
            If .width + X - sngMouseX > sngMinWidth Then
                .width = .width + X - sngMouseX
            Else
                X = 0
                sngMouseX = 0
            End If
            If .height + Y - sngMouseY > sngMinHeight Then
                .height = .height + Y - sngMouseY
            Else
                Y = 0
                sngMouseY = 0
            End If
        End With
        If X <> 0 Or Y <> 0 Then
            RaiseEvent Resizing(X - sngMouseX, Y - sngMouseY)
        End If
    End If
End Sub

Private Sub frResizerRight_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        With oResizableForm
            If .width + X > sngMinWidth Then
                .width = .width + X
                RaiseEvent Resizing(X, 0)
            End If
        End With
    End If
End Sub

Private Sub frResizerBottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        With oResizableForm
            If .height + Y > sngMinHeight Then
                .height = .height + Y
                RaiseEvent Resizing(0, Y)
            End If
        End With
    End If
End Sub

Private Sub frmResizableForm_Layout()
    With frResizerCorner
        .Left = oResizableForm.InsideWidth - .width
        .Top = oResizableForm.InsideHeight - .height
    End With
    With frResizerRight
        .Left = oResizableForm.InsideWidth - .width
        .height = frResizerCorner.Top
    End With
    With frResizerBottom
        .Top = oResizableForm.InsideHeight - .height
        .width = frResizerCorner.Left
    End With
End Sub
