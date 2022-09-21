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
   TypeInfoVer     =   105
End
Attribute VB_Name = "PlantUMLEdit"
Attribute VB_Base = "0{17C463E1-DDBC-4909-9F38-832D32AA2A81}{E0193FC7-C9E4-49DD-89A6-0C928B3CF82B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Public Parent As Object
Public Hidden As Boolean
Private Focus As Boolean
Private Target As Shape
Private WithEvents oFormResize As UserFormResizer
Attribute oFormResize.VB_VarHelpID = -1
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private Initializing As Boolean


Private Sub App_WindowDeactivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
    Hide
End Sub

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
On Error GoTo Failed
    If Not Parent Is ActiveWindow Then
        Exit Sub
    End If
    If Not Hidden And ActiveWindow.Selection.Type = ppSelectionShapes _
            And ActiveWindow.Selection.ShapeRange.Count = 1 _
            And ActiveWindow.Selection.ShapeRange(1).Tags("diagram_type") > "" Then

        If Left <> 0 Then
            StartUpPosition = 0
        End If
        
        ShowWindow Focus
        Focus = False
        
        Exit Sub
    End If
Failed:
    Hide
End Sub

Private Sub BrowseForJarButton_Click()
    PlantUml.BrowseForJar
    JarLocationTextBox.Text = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
End Sub

Private Sub UpdateDiagram(Optional Force As Boolean = False)
    If Initializing Then
        Exit Sub
    End If
    WorkingLabel.Caption = "Working..."
    Dim continue As Boolean
    Do
        continue = PlantUml.UpdateDiagram(Target, Code.Text, TypeCombo.Text, Force)
        DoEvents
    Loop While continue And Not Force
    WorkingLabel.Caption = ""
End Sub

Private Sub CancelButton_Click()
    Hidden = True
    Hide
End Sub

Private Sub Code_Change()
    UpdateDiagram
End Sub


Private Sub FormatCombo_Change()
    SaveSetting "PlantUML_Plugin", "Settings", "Format", FormatCombo.Text
    UpdateDiagram True
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
    Initializing = True
    
    On Error Resume Next
    Set Target = ActiveWindow.Selection.ShapeRange(1)
    TypeCombo.Text = Target.Tags("diagram_type")
    Code.Text = Target.Tags("plantuml")
    Code.SelStart = 0
    
    Initializing = False
End Sub


Private Sub UserForm_Initialize()
    
    Initializing = True
    Set App = Application
    Set Parent = ActiveWindow
    
    TypeCombo.AddItem "uml"
    TypeCombo.AddItem "gantt"
    TypeCombo.AddItem "mindmap"
    TypeCombo.AddItem "wbs"
    
    FormatCombo.AddItem "svg"
    FormatCombo.AddItem "png"
    
    JarLocationTextBox.Text = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
    FormatCombo.Text = GetSetting("PlantUML_Plugin", "Settings", "Format", "svg")
    
    Set oFormResize = New UserFormResizer
    Set oFormResize.ResizableForm = Me
    
    If Dir(JarLocationTextBox.Text) = "" Then
        BrowseForJarButton_Click
    End If
    
    
End Sub

Private Sub oFormResize_Resizing(ByVal X As Single, ByVal Y As Single)
    Dim Control As Control
    For Each Control In Controls
        Dim Tag
        With Control
            For Each Tag In Split(.Tag, ",")
                Select Case Tag
                Case "width"
                    .Width = .Width + X
                Case "height"
                    .Height = .Height + Y
                Case "bottom"
                    .Top = .Top + Y
                Case "right"
                    .Left = .Left + X
                End Select
            Next
        End With
    Next
End Sub

Private Sub AlignBottom(ctl As Control, ByVal Y As Single)
    ctl.Top = ctl.Top + Y
End Sub

Private Sub AlignRight(ctl As Control, ByVal X As Single)
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

Private Sub ShowWindow(Optional Focus As Boolean = True)
    UserForm_Activate
        
    Show
    
    TypeCombo.SetFocus
    Code.SetFocus
    
    If Not Focus Then
        Target.Select
    End If
End Sub

Public Sub Edit(Optional shp As Shape)
    
    If shp Is Nothing Then
        ShowWindow
    Else
        Focus = True
        Hidden = False
        shp.Select
    End If
End Sub
