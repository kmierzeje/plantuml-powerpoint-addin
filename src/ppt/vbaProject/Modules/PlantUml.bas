Attribute VB_Name = "PlantUml"
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
    
    shp.Select
    EditDiagram
    
    If shp.Tags("plantuml") = "" Then
        shp.Delete
    End If
End Sub

Public Sub EditDiagram()
    If ActiveWindow.Selection.ShapeRange.Count = 0 Then
        Exit Sub
    End If
    PlantUMLEdit.Show
End Sub

Function GetScale(orig As String, current As Single) As Single
    If orig = "" Then
        GetScale = 1#
    Else
        GetScale = current / Val(orig)
    End If

End Function


Public Function UpdateDiagram(shp As Shape, body As String, tag As String)
    On Error GoTo Failed
    UpdateDiagram = False
    
    
    If body = shp.Tags("plantuml") And shp.Tags("diagram_type") = tag Then
        Exit Function
    End If

    If body = "" Then
        shp.Fill.Transparency = 1#
        Exit Function
    End If
    UpdateDiagram = True
    
    Dim fname As String
    fname = GenerateDiagram(body, tag, "svg")
    
    shp.Tags.Add "plantuml", body
    shp.Tags.Add "diagram_type", tag
    
    SetPicture shp, fname
Failed:

End Function


Function Maximum(v1 As Single, v2 As Single)
    If v1 > v2 Then
        Maximum = v1
    Else
        Maximum = v2
    End If
End Function

Public Sub SetPicture(shp As Shape, fname As String)
    shp.Fill.UserPicture (fname)
    
    Set svg = CreateObject("Msxml2.DOMDocument")
    svg.Load fname
    Kill fname
    
    Dim w As Single, h As Single, scaleX As Single, scaleY As Single
    
    scaleX = GetScale(shp.Tags("orig_width"), shp.width)
    scaleY = GetScale(shp.Tags("orig_height"), shp.height)
    
    w = Val(svg.SelectSingleNode("/svg/@width").text)
    h = Val(svg.SelectSingleNode("/svg/@height").text)
    
    shp.Tags.Add "orig_width", w
    shp.Tags.Add "orig_height", h
    
    shp.width = w * scaleX
    shp.height = h * scaleY

End Sub

