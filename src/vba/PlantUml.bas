Attribute VB_Name = "PlantUml"
Private controller As New UIController
Private editors As Collection
Private PlantServer As Object

Function CreateEditor() As PlantUMLEdit
    Dim key As String
    key = str(ObjPtr(Application.ActiveWindow))
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

Sub SetJarPath(path As String)
    SaveSetting "PlantUML_Plugin", "Settings", "JarPath", path
End Sub

Function BrowseForJar()
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Title = "Path to plantuml.jar"
        .Filters.Clear
        .Filters.Add "Jar Files", "*.jar", 1
        .InitialFileName = GetJarPath(False)
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Function
        End If
        BrowseForJar = .SelectedItems(1)
        SetJarPath .SelectedItems(1)
    End With
End Function

Function Q(text As String)
    Q = """" & text & """"
End Function

Function StringToHex(text As String) As String
    Dim out As String
    Dim chars() As Byte
    chars = Utils.ToUTF8(text)
    For i = LBound(chars) To UBound(chars)
        Dim Ch As String
        Ch = Hex(chars(i))
        If Len(Ch) = 1 Then
            Ch = "0" & Ch
        End If
        out = out & Ch
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
    
Function GetPicowebEndpoint() As String
    If GetJarPath(False) > "" Then
        GetPicowebEndpoint = GetSetting("PlantUML_Plugin", "Settings", "PicowebEndpoint", "8880")
    End If
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
        GetHttpServerAddress = GetRemoteHttpAddress()
    End If
End Function

Function GetRemoteHttpAddress()
    GetRemoteHttpAddress = GetSetting("PlantUML_Plugin", "Settings", "HttpServerAddress")
    If GetRemoteHttpAddress = "" Then
        GetRemoteHttpAddress = "https://www.plantuml.com"
    End If
End Function

Function SetRemoteHttpAddress(address As String)
    SaveSetting "PlantUML_Plugin", "Settings", "HttpServerAddress", address
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

Function GenerateDiagram(request As String, format As String)
    If GetJarPath(False) > "" And GetPicowebEndpoint() = "" Then
        GenerateDiagram = GenerateDiagramCmd(request, format)
    Else
        GenerateDiagram = GenerateDiagramHttp(request, format)
    End If
End Function

Function GenerateDiagramHttp(request As String, format As String)
    StartServer
    
    Dim WinHttpReq As Object
    Set WinHttpReq = VBA.CreateObject("WinHttp.WinHttpRequest.5.1")
        
    WinHttpReq.Open "GET", GetHttpServerAddress() & "/plantuml/" & format & "/~h" & StringToHex(request), True
    WinHttpReq.Send
    WinHttpReq.WaitForResponse
    Dim response() As Byte
    response = WinHttpReq.ResponseBody
    GenerateDiagramHttp = WriteToTmpBinFile(response, format)
End Function

Function GenerateDiagramCmd(request As String, format As String)
    Dim fname As String
    
    fname = WriteToTmpFile(request)
    
    SyncShell "java.exe -jar " & Q(GetJarPath()) & " -t" & format & " " & Q(fname), vbHide
    Kill fname
    fname = Left(fname, InStrRev(fname, ".") - 1) & "." & format
    GenerateDiagramCmd = fname
End Function


Public Sub InsertDiagram()
    Dim sld As Slide
    Dim shp As Shape
    Set sld = Application.ActiveWindow.View.Slide
    
    With Application.ActivePresentation.PageSetup
        Set shp = sld.Shapes.AddShape(msoShapeRectangle, .SlideWidth / 4, .SlideHeight / 4, .SlideWidth / 2, .SlideHeight / 2)
    End With
    shp.Fill.Transparency = 1#
    With shp.TextFrame.TextRange.Font
        .Color = RGB(0, 0, 0)
        .size = 14
    End With
    
    shp.Line.Visible = msoFalse
    shp.Tags.Add "plantuml", ""
    shp.Tags.Add "diagram_type", "uml"
    shp.Tags.Add "scaling", 0
    
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

Public Sub RefreshDiagram(shp As Shape)
    On Error Resume Next
    If shp.Tags("diagram_type") = "" Or shp.Tags("scaling") <> "0" Then
        Exit Sub
    End If
    
    AdjustCropping shp
End Sub

Function EncodeColor(Number As Long) As String
    EncodeColor = Hex(Number)
    EncodeColor = String(6 - Len(EncodeColor), "0") & EncodeColor
    EncodeColor = Right(EncodeColor, 2) & Mid(EncodeColor, 3, 2) & Left(EncodeColor, 2)
End Function

Private Function UpdateTag(shp As Shape, name As String, value As Variant) As Boolean
    If shp.Tags(name) = value Then
        Exit Function
    End If
    UpdateTag = True
    shp.Tags.Add name, value
End Function

Private Function Clean(text As String) As String
    Dim i As Integer, final As Integer
    Dim ws As String
    
    
    For i = 1 To Len(text)
        If Mid(text, i, 6) = vbNewLine & "@end" Then
            Exit For
        End If
        Dim Ch As String
        Ch = Mid(text, i, 1)
        If InStr(" " & vbNewLine & vbTab, Ch) > 0 Then
            ws = ws & Ch
        ElseIf Ch <> vbCr Then
            If Ch = "\" Then
               If InStr("\rnt" & vbNewLine, Mid(text, i + 1, 1)) = 0 Then Clean = Clean + "\"
            End If
            Clean = Clean & ws & Ch
            ws = ""
        End If
    Next i
    
    If Clean = "" Then
        Exit Function
    ElseIf Mid(Clean, Len(Clean)) = "\" Then
        Clean = Clean & "\"
    End If
End Function


Public Function UpdateDiagram(shp As Shape, body As String, Tag As String, Theme As String, Scaling As Long, Optional Force As Boolean = False)
    On Error GoTo Failed
    UpdateDiagram = False
    
    Dim FontDecl As String
    'With shp.TextFrame.TextRange.Font
    '    FontDecl = "skinparam defaultFontName " & .Name & vbNewLine
    '    FontDecl = FontDecl & "skinparam defaultFontSize " & Str(.Size) & vbNewLine
    '    FontDecl = FontDecl & "skinparam defaultFontColor " & EncodeColor(.Color.RGB) & vbNewLine
    'End With
    Dim modified As Boolean
    
    modified = UpdateTag(shp, "plantuml", body) _
           Or UpdateTag(shp, "diagram_type", Tag) _
           Or UpdateTag(shp, "theme", Theme) _
           Or UpdateTag(shp, "font", FontDecl) _
           Or UpdateTag(shp, "scaling", Scaling) _
           Or Force
    
    If Not modified Then
        Exit Function
    End If
    
    body = Clean(body)

    If body = "" Then
        shp.Fill.Transparency = 1#
        Exit Function
    End If
    UpdateDiagram = True
    
    Dim ThemeDecl As String
    If Theme > "" Then
        ThemeDecl = "!theme " & Theme & vbNewLine
    End If
    
    Dim Code As String
    Code = "@start" & Tag & vbNewLine & FontDecl & ThemeDecl & body & vbNewLine & "@end" & Tag
    
    Dim format As String
    format = GetSetting("PlantUML_Plugin", "Settings", "Format")
    
    Dim fname As String
    fname = GenerateDiagram(Code, format)
    
    SetPicture shp, fname, format, Scaling
    Exit Function
Failed:
    MsgBox Err.Description, vbCritical, "PlantUml", Err.HelpFile, Err.HelpContext
End Function


Function Maximum(v1 As Single, v2 As Single)
    If v1 > v2 Then
        Maximum = v1
    Else
        Maximum = v2
    End If
End Function

Private Sub AdjustCropping(shp As Shape)
    Dim CropScale As Single, w As Single, h As Single
    w = Val(shp.Tags("orig_width"))
    h = Val(shp.Tags("orig_height"))
    
    CropScale = shp.width / w
    If shp.height / h < CropScale Then
        CropScale = shp.height / h
    End If
    With shp.PictureFormat.Crop
        .PictureWidth = w * CropScale
        .PictureHeight = h * CropScale
    End With

End Sub

Public Sub SetPicture(shp As Shape, fname As String, format As String, Scaling As Long)
    On Error Resume Next
    Dim scaleX As Single, scaleY As Single
    If shp.Fill.Type = msoFillPicture Then
        With shp.PictureFormat.Crop
            scaleX = GetScale(shp.Tags("orig_width"), .PictureWidth)
            scaleY = GetScale(shp.Tags("orig_height"), .PictureHeight)
        End With
    End If
    
    shp.Fill.UserPicture (fname)
    
    Dim w As Single, h As Single
    
    If format = "svg" Then
        Set svg = CreateObject("Msxml2.DOMDocument")
        svg.Load fname
    
        w = Val(svg.SelectSingleNode("/svg/@width").text)
        h = Val(svg.SelectSingleNode("/svg/@height").text)
    Else
        Set wia = CreateObject("WIA.ImageFile")
        wia.LoadFile fname
        w = wia.width
        h = wia.height
    End If
    
    
    shp.Tags.Add "orig_width", w
    shp.Tags.Add "orig_height", h
    
    Dim lockedAspect As MsoTriState
    lockedAspect = shp.LockAspectRatio
    shp.LockAspectRatio = msoFalse
    
    If Scaling = 1 Then
        shp.width = w * scaleX
        shp.height = h * scaleY
    Else
        AdjustCropping shp
    End If
    
    
    shp.LockAspectRatio = lockedAspect
    Kill fname
End Sub

Sub PlantUMLBtn_GetEnabled(Control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = Not Application.ActiveWindow.View.Slide Is Nothing
End Sub

Sub PlantUMLEdit_GetVisible(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetSelectedShape().Tags("diagram_type") > ""
End Sub

Public Function GetSelectedShape()
    On Error Resume Next
    Set GetSelectedShape = ActiveWindow.Selection.TextRange.Parent.Parent.Item(1)
End Function

