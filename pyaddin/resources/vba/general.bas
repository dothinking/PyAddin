Attribute VB_Name = "General"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GENERAL FUNCTIONS CREATED AUTOMATICALLY BY PYADDIN
'
' RunPython(python_method_name, arg1, arg2, ...) is a pre-defined VBA function to call
' Python scripts from command line, and check result from output/error file generated
' by the called Python script.
'
' https://github.com/dothinking/PyAddin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)


' global parameters
Public Const CFG_FILE = "main.cfg"
Public Const PYTHON_MAIN = "main.py"

Public Const KEY_STATUS = "status"
Public Const KEY_VALUE = "value"

Public Const KEY_PYTHON = "python"
Public Const KEY_OUTPUT = "output"
Public Const KEY_STDOUT = "stdout"
Public Const KEY_STDERR = "stderr"
Public Const KEY_RIBBON = "ribbon"

Public PYTHON_PATH As String
Public OUTPUT_PATH As String
Public STDOUT_PATH As String
Public STDERR_PATH As String


Function RunPython(methodName As String, ParamArray args()) As Object
    '''
    ' Run python script.
    ' :param methodName: a string refer to the called python method -> package.module.method.
    ' :param args: dynamic parameters for python arguments.
    ' :returns: result Dictionary with two keys: status and value.
    '
    '''
    
    ' check python interpreter path
    Dim errs As String
    If Not FileExists(PYTHON_PATH) Then
        errs = "Please set correct Python interpreter or refresh settings."
        GoTo OUTPUT
    End If
    
    ' join command: python main.py workBookName pythonMethod arguments
    Dim python As String: python = """" & PYTHON_PATH & """ "
    Dim main As String: main = """" & ThisWorkbook.Path & "\" & PYTHON_MAIN & """ "
    Dim wbName As String: wbName = """" & ActiveWorkbook.name & """ "
    
    Dim param, strArgs As String, cmd As String
    For Each param In args
        strArgs = strArgs & " """ & param & """"
    Next
    
    cmd = python & main & wbName & methodName & strArgs
    
    ' execute command
    Dim oShell As Object:
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run cmd, 0, 1
    
    ' results
    If Not FileExists(STDERR_PATH) Then
        errs = "Run Python script failed."
        GoTo OUTPUT
    End If
    
    errs = ReadTextFile(STDERR_PATH)
    
OUTPUT:
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")

    If errs = "" Then
        res.Add KEY_STATUS, True
        res.Add KEY_VALUE, ReadTextFile(STDOUT_PATH)
    Else:
        res.Add KEY_STATUS, False
        res.Add KEY_VALUE, errs
    End If
    
    Set RunPython = res
    
    Set oShell = Nothing ' clean object
    
    ' remove log file
    Call ClearOutput
    
End Function


Sub LoadConfig()
    '''
    ' get configuration data from main.cfg
    '
    '''
    PYTHON_PATH = GetConfig(KEY_PYTHON)
    OUTPUT_PATH = ThisWorkbook.Path & "\" & GetConfig(KEY_OUTPUT) & "\"
    STDOUT_PATH = OUTPUT_PATH & GetConfig(KEY_STDOUT)
    STDERR_PATH = OUTPUT_PATH & GetConfig(KEY_STDERR)
End Sub


Function GetConfig(sName As String, Optional sValue As String) As String
    '''
    ' get specified configuration data from main.cfg
    '
    '''
    GetConfig = sValue
    
    Dim cfg_path As String: cfg_path = ThisWorkbook.Path & "\" & CFG_FILE
    If Not FileExists(cfg_path) Then Exit Function
    
    Dim str As String
    Dim target As String: target = "[" & sName & "]"
        
    Open cfg_path For Input As #1
    Do While Not EOF(1)
        Line Input #1, str
        str = Trim(str)
        
        If str = target Then
            Line Input #1, str
            GetConfig = str
        End If
    Loop
    Close #1
End Function


Sub SetConfig(sName As String, Optional sValue As String)
    '''
    ' write configuration
    '
    '''
    Dim cfg_path As String: cfg_path = ThisWorkbook.Path & "\" & CFG_FILE
    If Not FileExists(cfg_path) Then Exit Sub
    
    ' collect parameters
    Dim target As String: target = "[" & sName & "]"
    Dim lines(1 To 100) As String
    Dim NUM As Integer, i As Integer
    Dim str As String
    
    Open cfg_path For Input As #1
    Do While Not EOF(1)
        Line Input #1, str
        str = Trim(str)
        
        NUM = NUM + 1
        lines(NUM) = str
        
        ' find target
        If str = target Then
            ' set new value
            NUM = NUM + 1
            lines(NUM) = sValue
            
            ' skip original line
            Line Input #1, str
        End If
    Loop
    Close #1
    
    ' write file
    Open cfg_path For Output As #2
    For i = 1 To NUM
        Print #2, lines(i)
    Next i
    Close #2

End Sub


Sub ClearOutput()
    '''
    ' remove output/error log files
    '
    '''
    If FileExists(STDOUT_PATH) Then Kill STDOUT_PATH
    If FileExists(STDERR_PATH) Then Kill STDERR_PATH
End Sub

Function FileExists(ByVal FileSpec As String) As Boolean
    '''
    ' check whether file exists
    '
    '''
   Dim Attr As Long
   On Error Resume Next
   Attr = GetAttr(FileSpec)
   If Err.Number = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      FileExists = Not ((Attr And vbDirectory) = vbDirectory)
   End If
End Function

Function ReadTextFile(filename As String) As String
    '''
    ' read content of text file
    '
    '''
    Dim str As String, txt As String
    Open filename For Input As #1
    Do While Not EOF(1)
        Line Input #1, str
        txt = txt & str
    Loop
    Close #1
    
    ReadTextFile = txt
    
End Function


Function RestoreRibbon()
    '''
    ' Restore Ribbon UI instance from stored pointer.
    '
    '''
    Dim p As Long: p = GetConfig(KEY_RIBBON)

    Dim ribbon As Object
    Call CopyMemory(ribbon, p, 4)
    
    Set RestoreRibbon = ribbon

End Function
