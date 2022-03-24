Attribute VB_Name = "PlantUml"
Private controller As New UIController

Function CreateEditor() As PlantUMLEdit
    Set CreateEditor = New PlantUMLEdit
End Function

Sub OnLoad(Ribbon As IRibbonUI)
    Set controller.Ribbon = Ribbon
End Sub

Sub SyncShell(ByVal Cmd As String, ByVal WindowStyle As VbAppWinStyle)
    VBA.CreateObject("WScript.Shell").Run Cmd, WindowStyle, True
End Sub

Function WriteToTmpFile(sText As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempFileName As String
    tempFileName = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
    
    Set ts = fso.CreateTextFile(tempFileName)
    ts.WriteLine sText
    ts.Close
    WriteToTmpFile = tempFileName
End Function

Sub BrowseForJar()
    With Application.FileDialog(msoFileDialogOpen)
            .AllowMultiSelect = False
            .Title = "Path to plantuml.jar"
            .Filters.Add "Jar Files", "*.jar", 1
            .InitialFileName = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
            .Show
            If .SelectedItems.Count = 0 Then
                Exit Sub
            End If
            
            SaveSetting "PlantUML_Plugin", "Settings", "JarPath", .SelectedItems(1)
        End With
End Sub

Function GenerateDiagram(body As String, tag As String, format As String)
    Dim fname As String
    
    fname = WriteToTmpFile("@start" & tag & vbNewLine & body & vbNewLine & "@end" & tag)
    
    Dim JarPath As String
    JarPath = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
    If JarPath = "" Then
        BrowseForJar
    End If
    
    SyncShell "java.exe -jar " & JarPath & " -t" & format & " " & fname, vbHide
    Kill fname
    fname = Left(fname, InStrRev(fname, ".") - 1) & "." & format
    GenerateDiagram = fname
End Function


Public Sub InsertDiagram()
    Dim sld As Slide
    Dim shp As Shape
    Set sld = Application.ActiveWindow.View.Slide
    
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, 0, 0, 1, 1)
    shp.Fill.Transparency = 1#
    shp.Line.Visible = msoFalse
    shp.Tags.Add "plantuml", ""
    shp.Tags.Add "diagram_type", "uml"
    
    CreateEditor().Edit shp
End Sub

Public Sub EditDiagram()
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        Exit Sub
    End If
    CreateEditor().Edit
End Sub

Function GetScale(orig As String, current As Single) As Single
    If orig = "" Then
        GetScale = 1#
    Else
        GetScale = current / Val(orig)
    End If

End Function


Public Function UpdateDiagram(shp As Shape, body As String, Tag As String, Optional Force As Boolean = False)
    On Error GoTo Failed
    UpdateDiagram = False
    
    body = Replace(body, vbCr, "")
    
    If Not Force And body = shp.Tags("plantuml") And shp.Tags("diagram_type") = Tag Then
        Exit Function
    End If
    
    shp.Tags.Add "plantuml", body
    shp.Tags.Add "diagram_type", tag

    If body = "" Then
        shp.Fill.Transparency = 1#
        Exit Function
    End If
    UpdateDiagram = True
    
    Dim fname As String
    
    Dim format As String
    format = GetSetting("PlantUML_Plugin", "Settings", "Format")
    fname = GenerateDiagram(body, Tag, format)
    
    SetPicture shp, fname, format
Failed:

End Function


Function Maximum(v1 As Single, v2 As Single)
    If v1 > v2 Then
        Maximum = v1
    Else
        Maximum = v2
    End If
End Function

Public Sub SetPicture(shp As Shape, fname As String, format As String)
    shp.Fill.UserPicture (fname)
    
    Dim w As Single, h As Single, scaleX As Single, scaleY As Single
    scaleX = GetScale(shp.Tags("orig_width"), shp.width)
    scaleY = GetScale(shp.Tags("orig_height"), shp.height)
    
    If format = "svg" Then
        Set svg = CreateObject("Msxml2.DOMDocument")
        svg.Load fname
    
        w = Val(svg.SelectSingleNode("/svg/@width").Text)
        h = Val(svg.SelectSingleNode("/svg/@height").Text)
    Else
        Set wia = CreateObject("WIA.ImageFile")
        wia.LoadFile fname
        w = wia.Width
        h = wia.Height
    End If
    
    
    shp.Tags.Add "orig_width", w
    shp.Tags.Add "orig_height", h
    
    shp.width = w * scaleX
    shp.height = h * scaleY
    
    Kill fname
End Sub

Sub PlantUMLBtn_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = Not Application.ActiveWindow.View.Slide Is Nothing
End Sub

Sub PlantUMLEdit_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange(1).Tags("diagram_type") > ""
End Sub
