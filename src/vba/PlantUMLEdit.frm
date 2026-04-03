VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlantUMLEdit 
   Caption         =   "PlantUML"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11412
   OleObjectBlob   =   "PlantUMLEdit.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlantUMLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Parent As Object
Public Hidden As Boolean
Private Focus As Boolean
Private Target As Shape
Private WithEvents oFormResize As UserFormResizer
Attribute oFormResize.VB_VarHelpID = -1
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private Initializing As Boolean


Private Sub App_PresentationCloseFinal(ByVal Pres As Presentation)
    On Error Resume Next
    PlantUml.StopServer
End Sub

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

Private Sub UpdateDiagram(Optional Force As Boolean = False)
    If Initializing Then
        Exit Sub
    End If
    WorkingLabel.Caption = "Working..."
    Dim continue As Boolean
    Do
        continue = PlantUml.UpdateDiagram(Target, Code.Text, TypeCombo.Text, ThemeCombo.Text, ScalingStyleCombo.ListIndex, Force)
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

Private Sub Code_KeyDown(pKey As Long, ByVal ShiftKey As Integer)
    Select Case pKey
    Case Asc(vbTab)
        Dim Indent As String: Indent = "  "
        If ShiftKey Then
            Outdent (Indent)
        Else
            If InStr(Code.SelText, Chr(13)) = 0 Then
                Code.SelText = Indent
                Code.SelStart = Code.SelStart + Code.SelLength
            Else
                IndentCode Indent
            End If
        End If
        pKey = 0
    Case vbKeyEscape
        pKey = 0
        CancelButton_Click
    End Select
End Sub

Private Sub FormatCombo_Change()
    SaveSetting "PlantUML_Plugin", "Settings", "Format", FormatCombo.Text
    UpdateDiagram True
End Sub

Private Sub ExportButton_Click()
    Dim FileName As String
    Dim Filters(4) As String
    Filters(0) = "PlantUML Files"
    Filters(1) = "*.puml"
    Filters(2) = "All Files"
    Filters(3) = "*.*"

    FileName = Utils.GetSaveFileNameWinAPI(Filters, "Export PlantUML", "diagram.puml")
    If FileName = "" Then
        Exit Sub
    End If
    Utils.WriteUTF8File FileName, "@start" & TypeCombo.Text & vbCrLf & Me.Code.Text & vbCrLf & "@end" & TypeCombo.Text
End Sub

Private Sub ImportButton_Click()
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Title = "Import PlantUML"
        .Filters.Clear
        .Filters.Add "PlantUML Files", "*.puml", 1
        .Filters.Add "All Files", "*.*", 2
        .InitialFileName = GetJarPath(False)
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If

        Dim Text As String
        Text = Utils.ReadUTF8File(.SelectedItems(1))
        Dim pos As Integer
        If Left(Text, 6) = "@start" Then
            pos = 1
        Else
            pos = InStr(Text, vbNewLine & "@start")
        End If
        If pos > 0 Then
            pos = InStr(pos + 1, Text, vbNewLine) + Len(vbNewLine)
            Text = Mid(Text, pos)
        End If
        
        pos = InStr(Text, vbNewLine & "@end")
        If pos > 1 Then
            While Mid(Text, pos - 1, 1) = "\"
                pos = InStr(pos + 1, Text, vbNewLine & "@end")
            Wend
            Text = Left(Text, pos - 1)
        End If
        Code.Text = Text
        Code.SetFocus
        Code.SelStart = 0
    End With
    
End Sub

Private Sub ScalingStyleCombo_Change()
    Code_Change
End Sub

Private Sub ServerComboBox_Change()
    If Initializing Then
        Exit Sub
    End If
    
    If ServerComboBox.ListIndex = -1 Then
        PlantUml.SetRemoteHttpAddress ServerComboBox.value
        Exit Sub
    ElseIf ServerComboBox.ListIndex = 0 Then
        PlantUml.SetRemoteHttpAddress ServerComboBox.value
        PlantUml.SetJarPath ""
    ElseIf ServerComboBox.ListIndex < ServerComboBox.ListCount - 1 Then
        PlantUml.SetJarPath ServerComboBox.value
    Else
        PlantUml.BrowseForJar
    End If
    
    SetupServerCombo
End Sub

Private Sub ThemeCombo_Change()
    Code_Change
End Sub

Private Sub TypeCombo_Change()
    EndLabel.Caption = "@end" & TypeCombo.Text
    Code_Change
End Sub

Private Sub SetupServerCombo()
    Initializing = True
    Dim LocalJarPath As String
    
    ServerComboBox.Clear
    ServerComboBox.AddItem PlantUml.GetRemoteHttpAddress()
    
    LocalJarPath = PlantUml.GetJarPath(False)
    If LocalJarPath > "" Then
        ServerComboBox.AddItem LocalJarPath
        ServerComboBox.value = LocalJarPath
        ServerComboBox.Style = fmStyleDropDownList
    Else
        ServerComboBox.value = PlantUml.GetRemoteHttpAddress()
        ServerComboBox.Style = fmStyleDropDownCombo
    End If
    ServerComboBox.AddItem "Browse for 'plantuml.jar'..."
    
    MeasureTextBox.Text = ServerComboBox.value
    If MeasureTextBox.Width > ServerComboBox.Width - 16 Then
        ServerComboBox.ControlTipText = ServerComboBox.value
    Else
        ServerComboBox.ControlTipText = ""
    End If
    Initializing = False
End Sub

Private Sub UserForm_Activate()
    Hidden = False
    SetupServerCombo
    Initializing = True
    FormatCombo.Text = GetSetting("PlantUML_Plugin", "Settings", "Format", "svg")
    
    On Error Resume Next
    Set Target = PlantUml.GetSelectedShape()
    TypeCombo.Text = Target.Tags("diagram_type")
    ThemeCombo.Text = Target.Tags("theme")
    If Target.Tags("scaling") = "" Then
        ScalingStyleCombo.ListIndex = 1
    Else
        ScalingStyleCombo.ListIndex = Target.Tags("scaling")
    End If
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
    
    ThemeCombo.AddItem ""
    ThemeCombo.AddItem "amiga"
    ThemeCombo.AddItem "aws-orange"
    ThemeCombo.AddItem "black-knight"
    ThemeCombo.AddItem "bluegray"
    ThemeCombo.AddItem "blueprint"
    ThemeCombo.AddItem "carbon-gray"
    ThemeCombo.AddItem "cerulean"
    ThemeCombo.AddItem "cerulean-outline"
    ThemeCombo.AddItem "cloudscape-design"
    ThemeCombo.AddItem "crt-amber"
    ThemeCombo.AddItem "crt-green"
    ThemeCombo.AddItem "cyborg"
    ThemeCombo.AddItem "cyborg-outline"
    ThemeCombo.AddItem "hacker"
    ThemeCombo.AddItem "lightgray"
    ThemeCombo.AddItem "mars"
    ThemeCombo.AddItem "materia"
    ThemeCombo.AddItem "materia-outline"
    ThemeCombo.AddItem "metal"
    ThemeCombo.AddItem "mimeograph"
    ThemeCombo.AddItem "minty"
    ThemeCombo.AddItem "mono"
    ThemeCombo.AddItem "plain"
    ThemeCombo.AddItem "reddress-darkblue"
    ThemeCombo.AddItem "reddress-darkgreen"
    ThemeCombo.AddItem "reddress-darkorange"
    ThemeCombo.AddItem "reddress-darkred"
    ThemeCombo.AddItem "reddress-lightblue"
    ThemeCombo.AddItem "reddress-lightgreen"
    ThemeCombo.AddItem "reddress-lightorange"
    ThemeCombo.AddItem "reddress-lightred"
    ThemeCombo.AddItem "sandstone"
    ThemeCombo.AddItem "silver"
    ThemeCombo.AddItem "sketchy"
    ThemeCombo.AddItem "sketchy-outline"
    ThemeCombo.AddItem "spacelab"
    ThemeCombo.AddItem "spacelab-white"
    ThemeCombo.AddItem "sunlust"
    ThemeCombo.AddItem "superhero"
    ThemeCombo.AddItem "superhero-outline"
    ThemeCombo.AddItem "toy"
    ThemeCombo.AddItem "united"
    ThemeCombo.AddItem "vibrant"
    
    ScalingStyleCombo.AddItem "diagram scale to shape size"
    ScalingStyleCombo.AddItem "shape size to diagram scale"
    ScalingStyleCombo.ListIndex = 0
    
    Set oFormResize = New UserFormResizer
    Set oFormResize.ResizableForm = Me
    
    PlantUml.StartServer
    
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

Private Sub SelectLines()
    Dim SelEnd As Integer: SelEnd = Code.SelStart + Code.SelLength
    
    While Code.SelStart > 0 And Code.SelText <> Chr(13)
        Code.SelStart = Code.SelStart - 1
        Code.SelLength = 1
    Wend
    If Code.SelStart > 1 Then
        Code.SelStart = Code.SelStart + 1
    End If
    Code.SelLength = SelEnd - Code.SelStart
    While Right(Code.SelText, 1) <> Chr(13)
        Dim SelLength As Integer: SelLength = Code.SelLength
        Code.SelLength = Code.SelLength + 1
        If Code.SelLength = SelLength Then
            Exit Sub
        End If
    Wend
    Code.SelLength = Code.SelLength - 1
End Sub

Function IndentCode(Indent As String) As Boolean
    SelectLines
    Dim Text As String: Text = Code.SelText
    
    Dim pos As Integer: pos = 0
    Do
        Text = Left(Text, pos) & Indent & Mid(Text, pos + 1)
        IndentCode = True
        pos = InStr(pos + Len(Indent), Text, Chr(13))
    Loop While pos > 0
    
    Dim SelStart As Integer: SelStart = Code.SelStart
    Code.SelText = Text
    Code.SelStart = SelStart
    Code.SelLength = Len(Text)
End Function

Private Function CommonPrefixLen(ByRef text1 As String, pos As Integer, text2 As String)
    Dim i As Integer
    For i = 1 To Len(text2)
        If Mid(text1, pos, i) <> Left(text2, i) Then
            Exit For
        End If
    Next i
    CommonPrefixLen = i - 1
End Function

Private Sub Outdent(Indent As String)
    SelectLines
    Dim Text As String: Text = Code.SelText
    
    Dim pos As Integer: pos = 0
    Do
        Dim size As Integer: size = CommonPrefixLen(Text, pos + 1, Indent)
        If size > 0 Then
            Text = Left(Text, pos) & Mid(Text, pos + size + 1)
        End If
        pos = InStr(pos + 1, Text, Chr(13))
    Loop While pos > 0
    
    Dim SelStart As Integer: SelStart = Code.SelStart
    Code.SelText = Text
    Code.SelStart = SelStart
    Code.SelLength = Len(Text)
End Sub

Private Sub Code_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
    Case Asc(vbTab)
        Dim Indent As String: Indent = "  "
        If Shift Then
            Outdent (Indent)
            KeyCode = 0
        Else
            If InStr(Code.SelText, Chr(13)) = 0 Then
                Code.SelText = Indent
                Code.SelStart = Code.SelStart + Code.SelLength
            Else
                IndentCode Indent
            End If
            KeyCode = 0
        End If
    Case vbKeyEscape
        KeyCode = 0
        CancelButton_Click
    End Select
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

