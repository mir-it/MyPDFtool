' Launch pdf_splitter_app.py with NO console window.
' Finds Python the same way as CMD: "where pythonw" / "where python" (PATH + App Paths).

Option Explicit

Dim sh, fs, folder, scriptPy, pyExe, rc, cmd, baseName

Set sh = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
folder = fs.GetParentFolderName(WScript.ScriptFullName)
scriptPy = folder & "\launch_gui.py"

If Not fs.FileExists(scriptPy) Then
  MsgBox "File not found:" & vbCrLf & scriptPy, vbCritical, "PDF Tools"
  WScript.Quit 1
End If
If Not fs.FileExists(folder & "\pdf_splitter_app.py") Then
  MsgBox "Missing pdf_splitter_app.py in:" & vbCrLf & folder, vbCritical, "PDF Tools"
  WScript.Quit 1
End If

sh.CurrentDirectory = folder

pyExe = FindPythonW(sh, fs)
If pyExe = "" Then
  MsgBox "Python was not found (pythonw.exe / pyw.exe)." & vbCrLf & vbCrLf & _
         "Your CMD works because PATH is set there; double-click uses a shorter PATH." & vbCrLf & vbCrLf & _
         "Try:" & vbCrLf & _
         "1) Run run_pdf_tools.bat" & vbCrLf & _
         "2) Reinstall Python from python.org and enable ""Add python.exe to PATH""" & vbCrLf & _
         "3) Or run in CMD: py -0p  (see install dir) and tell pythonw.exe path exists", _
         vbExclamation, "PDF Tools"
  WScript.Quit 1
End If

baseName = LCase(Mid(pyExe, InStrRev(pyExe, "\") + 1))
If baseName = "pyw.exe" Then
  cmd = """" & pyExe & """ -3 """ & scriptPy & """"
Else
  cmd = """" & pyExe & """ """ & scriptPy & """"
End If

rc = sh.Run(cmd, 0, True)
If rc <> 0 Then
  MsgBox "Python exited with code " & rc & "." & vbCrLf & vbCrLf & _
         "Details were saved to:" & vbCrLf & folder & "\data\last_startup_error.txt" & vbCrLf & vbCrLf & _
         "launch_gui.py should auto-install PyMuPDF; if it still fails, open that file or run run_pdf_tools.bat.", _
         vbExclamation, "Startup error"
End If

' --- Same discovery order as a normal CMD session (where uses PATH + registry App Paths) ---
Function WhereFirstExe(sh0, fs0, exeName)
  Dim tmp, comspec, ts, line
  tmp = sh0.ExpandEnvironmentStrings("%TEMP%\pdf_tools_where_out.txt")
  comspec = sh0.ExpandEnvironmentStrings("%ComSpec%")
  On Error Resume Next
  If fs0.FileExists(tmp) Then fs0.DeleteFile tmp
  Err.Clear
  sh0.Run comspec & " /c where """ & exeName & """ > """ & tmp & """ 2>nul", 0, True
  WhereFirstExe = ""
  If Not fs0.FileExists(tmp) Then Exit Function
  On Error Resume Next
  Set ts = fs0.OpenTextFile(tmp, 1, False, 0)
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  If Not ts.AtEndOfStream Then
    line = Trim(ts.ReadLine)
    If Len(line) > 0 Then
      If fs0.FileExists(line) Then WhereFirstExe = line
    End If
  End If
  ts.Close
  On Error Resume Next
  fs0.DeleteFile tmp
  On Error GoTo 0
End Function

Function PythonwBesidePython(fs0, pythonPath)
  Dim i, d, cand
  PythonwBesidePython = ""
  i = InStrRev(pythonPath, "\")
  If i < 1 Then Exit Function
  d = Left(pythonPath, i)
  cand = d & "pythonw.exe"
  If fs0.FileExists(cand) Then PythonwBesidePython = cand
End Function

Function FindPythonW(sh0, fs0)
  Dim p, base, fld, sf, v, vers, i

  ' 1) PATH (matches ""where pythonw"" in CMD)
  p = WhereFirstExe(sh0, fs0, "pythonw.exe")
  If Len(p) > 0 Then
    FindPythonW = p
    Exit Function
  End If

  p = WhereFirstExe(sh0, fs0, "pyw.exe")
  If Len(p) > 0 Then
    FindPythonW = p
    Exit Function
  End If

  ' 2) python.exe on PATH -> pythonw.exe in same folder (some setups only add python)
  p = WhereFirstExe(sh0, fs0, "python.exe")
  If Len(p) > 0 Then
    p = PythonwBesidePython(fs0, p)
    If Len(p) > 0 Then
      FindPythonW = p
      Exit Function
    End If
  End If

  ' 3) py launcher
  p = sh0.ExpandEnvironmentStrings("%SystemRoot%\pyw.exe")
  If fs0.FileExists(p) Then
    FindPythonW = p
    Exit Function
  End If

  ' 4) Registry (python.org installer)
  vers = Array("3.14", "3.13", "3.12", "3.11", "3.10", "3.9", "3.8")
  For i = 0 To UBound(vers)
    v = vers(i)
    On Error Resume Next
    p = sh0.RegRead("HKCU\Software\Python\PythonCore\" & v & "\InstallPath\")
    If Err.Number = 0 Then
      If Len(p) > 0 Then
        If Right(p, 1) <> "\" Then p = p & "\"
        If fs0.FileExists(p & "pythonw.exe") Then
          FindPythonW = p & "pythonw.exe"
          Exit Function
        End If
      End If
    End If
    Err.Clear
    p = sh0.RegRead("HKLM\SOFTWARE\Python\PythonCore\" & v & "\InstallPath\")
    If Err.Number = 0 Then
      If Len(p) > 0 Then
        If Right(p, 1) <> "\" Then p = p & "\"
        If fs0.FileExists(p & "pythonw.exe") Then
          FindPythonW = p & "pythonw.exe"
          Exit Function
        End If
      End If
    End If
    Err.Clear
  Next

  ' 5) Per-user install folder
  base = sh0.ExpandEnvironmentStrings("%LocalAppData%\Programs\Python\")
  If fs0.FolderExists(base) Then
    Set fld = fs0.GetFolder(base)
    For Each sf In fld.SubFolders
      p = sf.Path & "\pythonw.exe"
      If fs0.FileExists(p) Then
        FindPythonW = p
        Exit Function
      End If
    Next
  End If

  ' 6) Program Files
  base = sh0.ExpandEnvironmentStrings("%ProgramFiles%\")
  If fs0.FolderExists(base) Then
    Set fld = fs0.GetFolder(base)
    For Each sf In fld.SubFolders
      If LCase(Left(sf.Name, 6)) = "python" Then
        p = sf.Path & "\pythonw.exe"
        If fs0.FileExists(p) Then
          FindPythonW = p
          Exit Function
        End If
      End If
    Next
  End If

  FindPythonW = ""
End Function
