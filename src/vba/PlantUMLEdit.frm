VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlantUMLEdit 
   OleObjectBlob   =   "PlantUMLEdit.frx":0000
   Caption         =   "PlantUML"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   93
End
Attribute VB_Name = "PlantUMLEdit"
Attribute VB_Base = "0{4D81A2E8-D919-48B6-85F7-481A6429260D}{D82CBFC3-4044-4614-8FF9-25C1420FC4F8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Hidden As Boolean

Private Target As Shape
Private WithEvents oFormResize As UserFormResizer
Attribute oFormResize.VB_VarHelpID = -1
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private Initializing As Boolean


Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
On Error GoTo Failed
    If Not Hidden And ActiveWindow.Selection.Type = ppSelectionShapes _
            And ActiveWindow.Selection.ShapeRange.Count = 1 _
            And ActiveWindow.Selection.ShapeRange(1).Tags("diagram_type") > "" Then

        UserForm_Initialize
        If Left <> 0 Then
            StartUpPosition = 0
        End If
        Show
        Target.Select
        Exit Sub
    End If
Failed:
    Hide
End Sub

Private Sub BrowseForJarButton_Click()
    PlantUml.BrowseForJar
    JarLocationTextBox.Text = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
End Sub

Private Sub UpdateDiagram()
    If Initializing Then
        Exit Sub
    End If
    WorkingLabel.Caption = "Working..."
    Dim continue As Boolean
    Do
        continue = PlantUml.UpdateDiagram(Target, Code.Text, TypeCombo.Text)
        DoEvents
    Loop While continue
    WorkingLabel.Caption = ""
End Sub

Private Sub CancelButton_Click()
    Hidden = True
    Hide
End Sub

Private Sub Code_Change()
    UpdateDiagram
End Sub

Private Sub JarLocationTextBox_Enter()
    BrowseForJarButton.SetFocus
    BrowseForJarButton_Click
End Sub

Private Sub TypeCombo_Change()
    EndLabel.Caption = "@end" & TypeCombo.Text
    Code_Change
End Sub

Private Sub UserForm_Activate()
    Hidden = False
End Sub

Private Sub UserForm_Initialize()
    Initializing = True
    
    Set App = Application
    
    TypeCombo.AddItem "uml"
    TypeCombo.AddItem "gantt"
    TypeCombo.AddItem "mindmap"
    TypeCombo.AddItem "wbs"
    
    JarLocationTextBox.Text = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
    
    If oFormResize Is Nothing Then
        Set oFormResize = New UserFormResizer
        Set oFormResize.ResizableForm = Me
    End If
    
    If Dir(JarLocationTextBox.Text) = "" Then
        BrowseForJarButton_Click
    End If
    
On Error GoTo Failed
    Set Target = ActiveWindow.Selection.ShapeRange(1)
    TypeCombo.Text = Target.Tags("diagram_type")
    Code.Text = Target.Tags("plantuml")
    Code.SelStart = 0
    Code.SetFocus
Failed:
    Initializing = False
End Sub

Private Sub oFormResize_Resizing(ByVal X As Single, ByVal Y As Single)
    With Code
        .width = .width + X
        .height = .height + Y
    End With
    
    AlignBottom JarLocationLabel, Y
    AlignBottom JarLocationTextBox, Y
    AlignBottom BrowseForJarButton, Y
    AlignBottom EndLabel, Y
    
End Sub

Private Sub AlignBottom(ctl As control, ByVal Y As Single)
    ctl.Top = ctl.Top + Y
End Sub

Private Sub AlignRight(ctl As control, ByVal X As Single)
    ctl.Left = ctl.Left + X
End Sub


Private Sub Code_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        CancelButton_Click
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> vbFormCode Then
        Cancel = 1
        CancelButton_Click
    End If
End Sub
