VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlantUMLEdit 
   Caption         =   "PlantUML"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   OleObjectBlob   =   "PlantUMLEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlantUMLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents oFormResize As UserFormResizer
Attribute oFormResize.VB_VarHelpID = -1
Private Initializing As Boolean

Private Sub BrowseForJarButton_Click()
    PlantUml.BrowseForJar
    JarLocationTextBox.text = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
End Sub

Private Sub UpdateDiagram()
    If Initializing Then
        Exit Sub
    End If
    
    Dim shp As Shape
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    Dim continue As Boolean
    Do
        continue = PlantUml.UpdateDiagram(shp, Code.text, TypeCombo.text)
        DoEvents
    Loop While continue
End Sub

Private Sub Code_Change()
    If Not LiveUpdatesCheckBox.Value Then
        Exit Sub
    End If
    UpdateDiagram
    
End Sub

Private Sub JarLocationTextBox_Enter()
    BrowseForJarButton.SetFocus
    BrowseForJarButton_Click
End Sub

Private Sub LiveUpdatesCheckBox_Change()
    Code_Change
End Sub

Private Sub TypeCombo_Change()
    EndLabel.Caption = "@end" & TypeCombo.text
    Code_Change
End Sub

Private Sub UserForm_Initialize()
    Initializing = True
    TypeCombo.AddItem "uml"
    TypeCombo.AddItem "gantt"
    TypeCombo.AddItem "mindmap"
    TypeCombo.AddItem "wbs"
    
    JarLocationTextBox.text = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
    LiveUpdatesCheckBox.Value = CBool(GetSetting("PlantUML_Plugin", "Settings", "LiveUpdates", "True"))
    
    Set oFormResize = New UserFormResizer
    Set oFormResize.ResizableForm = Me
    If Dir(JarLocationTextBox.text) = "" Then
        BrowseForJarButton_Click
    End If
    
    Dim shp As Shape
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    TypeCombo.text = shp.Tags("diagram_type")
    Code.text = shp.Tags("plantuml")
    Code.SelStart = 0
    
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
    AlignBottom LiveUpdatesCheckBox, Y
    AlignRight LiveUpdatesCheckBox, X
    
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
        Unload Me
    End If
End Sub

Private Sub UserForm_Terminate()
    SaveSetting "PlantUML_Plugin", "Settings", "LiveUpdates", CStr(LiveUpdatesCheckBox.Value)
    UpdateDiagram
End Sub
