Attribute VB_Name = "modCamConfig"
Option Explicit

' Настройки в текстовом INI: %APPDATA%\CorelCamMill\mill.ini
' Формат: key=value, строки с ";" и "[секция]" игнорируются.

Public Sub CfgDefaults(ByRef c As CamSettings)
    c.ToolName = "Flat 6 mm"
    c.ToolDiameter = 6#
    c.ToolNumber = 1
    c.TipAngle = 90#
    c.FeedXY = 1200#
    c.FeedZ = 300#
    c.SpindleRPM = 18000
    c.SafeZ = 20#
    c.StartZ = 3#
    c.StepDown = 2#
    c.TotalDepth = 6#
    c.StepoverPercent = 45#
    c.Mode = "Contour"
    c.Side = "Auto"
    c.Climb = True
    c.ProgramNumber = 1001
    c.UseATC = True
    c.Coolant = True
    c.OutputPath = ""
    c.DrawPreview = True
    c.FlattenTol = 0.05
End Sub

Public Function CfgPath() As String
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then base = Environ$("USERPROFILE")
    CfgPath = base & "\CorelCamMill\mill.ini"
End Function

Public Sub CfgLoad(ByRef c As CamSettings)
    Dim text As String, lines() As String, i As Long
    Dim ln As String, p As Long, k As String, v As String

    CfgDefaults c
    If Not FileExists(CfgPath()) Then
        CfgSave c
        Exit Sub
    End If
    text = ReadTextFile(CfgPath())
    lines = Split(Replace(text, vbLf, ""), vbCr)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) <> ";" And Left$(ln, 1) <> "[" Then
                p = InStr(ln, "=")
                If p > 1 Then
                    k = LCase$(Trim$(Left$(ln, p - 1)))
                    v = Trim$(Mid$(ln, p + 1))
                    CfgApply c, k, v
                End If
            End If
        End If
    Next i
End Sub

Private Sub CfgApply(ByRef c As CamSettings, ByVal k As String, ByVal v As String)
    Select Case k
        Case "name": c.ToolName = v
        Case "diameter": c.ToolDiameter = Val(v)
        Case "toolnumber": c.ToolNumber = CLng(Val(v))
        Case "tipangle": c.TipAngle = Val(v)
        Case "feedxy": c.FeedXY = Val(v)
        Case "feedz": c.FeedZ = Val(v)
        Case "spindlerpm": c.SpindleRPM = CLng(Val(v))
        Case "safez": c.SafeZ = Val(v)
        Case "startz": c.StartZ = Val(v)
        Case "stepdown": c.StepDown = Val(v)
        Case "totaldepth": c.TotalDepth = Val(v)
        Case "stepoverpercent": c.StepoverPercent = Val(v)
        Case "mode": c.Mode = v
        Case "side": c.Side = v
        Case "climb": c.Climb = (Val(v) <> 0)
        Case "programnumber": c.ProgramNumber = CLng(Val(v))
        Case "useatc": c.UseATC = (Val(v) <> 0)
        Case "coolant": c.Coolant = (Val(v) <> 0)
        Case "outputpath": c.OutputPath = v
        Case "drawpreview": c.DrawPreview = (Val(v) <> 0)
        Case "flattentol": c.FlattenTol = Val(v)
    End Select
End Sub

Public Sub CfgSave(ByRef c As CamSettings)
    Dim s As String
    s = "; CorelCAM — параметры фрезеровки" & vbCrLf
    s = s & "; Mode: Contour | Pocket | Drill | VCarve" & vbCrLf
    s = s & "; Side: Auto | Outside | Inside" & vbCrLf
    s = s & "[Tool]" & vbCrLf
    s = s & "Name=" & c.ToolName & vbCrLf
    s = s & "Diameter=" & CamFmt(c.ToolDiameter, 3) & vbCrLf
    s = s & "ToolNumber=" & c.ToolNumber & vbCrLf
    s = s & "TipAngle=" & CamFmt(c.TipAngle, 1) & vbCrLf
    s = s & "[Cut]" & vbCrLf
    s = s & "FeedXY=" & CamFmt(c.FeedXY, 1) & vbCrLf
    s = s & "FeedZ=" & CamFmt(c.FeedZ, 1) & vbCrLf
    s = s & "SpindleRPM=" & c.SpindleRPM & vbCrLf
    s = s & "SafeZ=" & CamFmt(c.SafeZ, 3) & vbCrLf
    s = s & "StartZ=" & CamFmt(c.StartZ, 3) & vbCrLf
    s = s & "StepDown=" & CamFmt(c.StepDown, 3) & vbCrLf
    s = s & "TotalDepth=" & CamFmt(c.TotalDepth, 3) & vbCrLf
    s = s & "StepoverPercent=" & CamFmt(c.StepoverPercent, 1) & vbCrLf
    s = s & "[Strategy]" & vbCrLf
    s = s & "Mode=" & c.Mode & vbCrLf
    s = s & "Side=" & c.Side & vbCrLf
    s = s & "Climb=" & CStr(Abs(CLng(c.Climb))) & vbCrLf
    s = s & "[Post]" & vbCrLf
    s = s & "ProgramNumber=" & c.ProgramNumber & vbCrLf
    s = s & "UseATC=" & CStr(Abs(CLng(c.UseATC))) & vbCrLf
    s = s & "Coolant=" & CStr(Abs(CLng(c.Coolant))) & vbCrLf
    s = s & "[Output]" & vbCrLf
    s = s & "OutputPath=" & c.OutputPath & vbCrLf
    s = s & "DrawPreview=" & CStr(Abs(CLng(c.DrawPreview))) & vbCrLf
    s = s & "FlattenTol=" & CamFmt(c.FlattenTol, 3) & vbCrLf
    EnsureFolder Left$(CfgPath(), InStrRev(CfgPath(), "\") - 1)
    WriteTextFile CfgPath(), s
End Sub

' ---------------------------------------------------------------- числа и файлы

' Число для G-кода и INI: точка как разделитель независимо от локали.
Public Function CamFmt(ByVal v As Double, ByVal decimals As Long) As String
    Dim s As String
    If Abs(v) < 0.0005 Then v = 0#
    s = Format$(v, "0." & String$(decimals, "0"))
    s = Replace(s, ",", ".")
    If InStr(s, ".") > 0 Then
        Do While Right$(s, 1) = "0"
            s = Left$(s, Len(s) - 1)
        Loop
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    End If
    If s = "-0" Or Len(s) = 0 Then s = "0"
    CamFmt = s
End Function

Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function

Public Sub EnsureFolder(ByVal path As String)
    On Error Resume Next
    If Len(path) > 0 Then
        If Len(Dir$(path, vbDirectory)) = 0 Then MkDir path
    End If
    On Error GoTo 0
End Sub

Public Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer, s As String
    f = FreeFile
    Open path For Binary As #f
    s = String$(LOF(f), " ")
    Get #f, , s
    Close #f
    ReadTextFile = s
End Function

Public Sub WriteTextFile(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Binary As #f
    Put #f, , text
    Close #f
End Sub
