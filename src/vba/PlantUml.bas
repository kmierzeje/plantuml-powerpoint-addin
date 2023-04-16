Attribute VB_Name = "PlantUml"
Private controller As New UIController
Private editors As Collection
Private PlantServer As Object
Private Const CP_UTF8 As Long = 65001

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As Long) As Long

Private Function ToUTF8(Text As String)
    Dim lngResult As Long
    Dim UTF8() As Byte

    lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), Len(Text), 0, 0, 0, 0)
    If lngResult > 0 Then
        ReDim UTF8(lngResult - 1)
        WideCharToMultiByte CP_UTF8, 0, StrPtr(Text), Len(Text), VarPtr(UTF8(0)), lngResult, 0, 0
    End If
    ToUTF8 = UTF8
End Function

Function CreateEditor() As PlantUMLEdit
    Dim key As String
    key = Str(ObjPtr(Application.ActiveWindow))
    If editors Is Nothing Then
        Set editors = New Collection
    End If
    
    On Error GoTo CreateNew
    Set CreateEditor = editors.Item(key)
    Exit Function
CreateNew:
    Set CreateEditor = New PlantUMLEdit
    editors.Add CreateEditor, key
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

Function GetJarPath(Optional Interactive As Boolean = True)
    GetJarPath = GetSetting("PlantUML_Plugin", "Settings", "JarPath")
    If Interactive And GetJarPath = "" Then
        GetJarPath = BrowseForJar()
    End If
End Function

Function BrowseForJar()
    With Application.FileDialog(msoFileDialogOpen)
            .AllowMultiSelect = False
            .Title = "Path to plantuml.jar"
            .Filters.Add "Jar Files", "*.jar", 1
            .InitialFileName = GetJarPath(False)
            .Show
            If .SelectedItems.Count = 0 Then
                Exit Function
            End If
            BrowseForJar = .SelectedItems(1)
            SaveSetting "PlantUML_Plugin", "Settings", "JarPath", .SelectedItems(1)
        End With
End Function

Function Q(Text As String)
    Q = """" & Text & """"
End Function

Function StringToHex(Text As String) As String
    Dim out As String
    Dim chars() As Byte
    chars = ToUTF8(Text)
    For i = LBound(chars) To UBound(chars)
        Dim ch As String
        ch = hex(chars(i))
        If Len(ch) = 1 Then
            ch = "0" & ch
        End If
        out = out & ch
    Next
    StringToHex = out
End Function

Function WriteToTmpBinFile(content() As Byte, format As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim FileName As String
    FileName = fso.GetSpecialFolder(2) & "\" & fso.GetTempName() & "." & format
    Dim fileNo As Integer
    fileNo = FreeFile
    Open FileName For Binary Lock Read Write As #fileNo
    Put #fileNo, , content
    Close #fileNo
    WriteToTmpBinFile = FileName
End Function

Function GetPicowebEndpoint()
    GetPicowebEndpoint = GetSetting("PlantUML_Plugin", "Settings", "PicowebEndpoint")
End Function

Function GetPicowebAddress()
    Dim endpoint() As String
    endpoint = Split(GetPicowebEndpoint(), ":")
    If UBound(endpoint) = -1 Then
        GetPicowebAddress = ""
    ElseIf UBound(endpoint) = 0 Then
        GetPicowebAddress = "http://127.0.0.1:" & endpoint(0)
    Else
        GetPicowebAddress = "http://" & endpoint(1) & ":" & endpoint(0)
    End If
End Function

Function GetHttpServerAddress()
    GetHttpServerAddress = GetPicowebAddress()
    
    If GetHttpServerAddress = "" Then
        GetHttpServerAddress = GetSetting("PlantUML_Plugin", "Settings", "HttpServerAddress", "https://www.plantuml.com")
    End If
End Function

Public Sub StartServer()
    If Not PlantServer Is Nothing Or GetPicowebEndpoint() = "" Then
        Exit Sub
    End If

    Set PlantServer = VBA.CreateObject("WScript.Shell").Exec("javaw.exe -jar " & Q(GetJarPath()) & " -picoweb:" & GetPicowebEndpoint())
End Sub

Public Sub StopServer()
    If PlantServer Is Nothing Or GetSetting("PlantUML_Plugin", "Settings", "KeepServerAfterExit", "no") <> "no" Then
        Exit Sub
    End If

    PlantServer.Terminate
    Set PlantServer = Nothing
End Sub

Function GenerateDiagram(body As String, Tag As String, format As String)
    If GetJarPath(False) > "" And GetPicowebEndpoint() = "" Then
        GenerateDiagram = GenerateDiagramCmd(body, Tag, format)
    Else
        GenerateDiagram = GenerateDiagramHttp(body, Tag, format)
    End If
End Function

Function GenerateDiagramHttp(body As String, Tag As String, format As String)
    Dim request As String
    request = "@start" & Tag & vbNewLine & body & vbNewLine & "@end" & Tag
    
    StartServer
    
    Dim WinHttpReq As WinHttp.WinHttpRequest
    Set WinHttpReq = New WinHttpRequest
        
    WinHttpReq.Open "GET", GetHttpServerAddress() & "/plantuml/" & format & "/~h" & StringToHex(request), True
    WinHttpReq.Send
    WinHttpReq.WaitForResponse
    'While WinHttpReq.WaitForResponse(0) = False
    '    DoEvents
    'Wend
        
    GenerateDiagramHttp = WriteToTmpBinFile(WinHttpReq.ResponseBody, format)
End Function

Function GenerateDiagramCmd(body As String, Tag As String, format As String)
    Dim fname As String
    
    fname = WriteToTmpFile("@start" & Tag & vbNewLine & body & vbNewLine & "@end" & Tag)
    
    SyncShell "java.exe -jar " & Q(GetJarPath()) & " -t" & format & " " & Q(fname), vbHide
    Kill fname
    fname = Left(fname, InStrRev(fname, ".") - 1) & "." & format
    GenerateDiagramCmd = fname
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
    'On Error GoTo Failed
    UpdateDiagram = False
    
    body = Replace(body, vbCr, "")
    
    If Not Force And body = shp.Tags("plantuml") And shp.Tags("diagram_type") = Tag Then
        Exit Function
    End If
    
    shp.Tags.Add "plantuml", body
    shp.Tags.Add "diagram_type", Tag

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
        w = wia.width
        h = wia.height
    End If
    
    
    shp.Tags.Add "orig_width", w
    shp.Tags.Add "orig_height", h
    
    shp.width = w * scaleX
    shp.height = h * scaleY
    
    Kill fname
End Sub

Sub PlantUMLBtn_GetEnabled(Control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = Not Application.ActiveWindow.View.Slide Is Nothing
End Sub

Sub PlantUMLEdit_GetVisible(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = ActiveWindow.Selection.ShapeRange.Count = 1 And ActiveWindow.Selection.ShapeRange(1).Tags("diagram_type") > ""
End Sub
