Attribute VB_Name = "general"
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

' global parameters
Public Const CFG_FILE = "main.cfg"
Public Const PYTHON_MAIN = "main.py"

Public Const KEY_STATUS = "status"
Public Const KEY_VALUE = "value"

Public PYTHON_PATH As String
Public OUTPUT_PATH As String
Public STDOUT As String
Public STDERR As String


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
    If Not PYTHON_PATH Like "*.exe" Then PYTHON_PATH = PYTHON_PATH & ".exe"
    If Not FileExists(PYTHON_PATH) Then
        errs = "Please set correct Python interpreter."
        GoTo OUTPUT
    End If
    
    ' join command: python main.py workBookName pythonMethod arguments
    Dim python As String: python = """" & PYTHON_PATH & """ "
    Dim main As String: main = """" & ThisWorkbook.Path & "\" & PYTHON_MAIN & """ "
    Dim wbName As String: wbName = """" & ActiveWorkbook.Name & """ "
    
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
    Dim logOutput As String: logOutput = OUTPUT_PATH & STDOUT
    Dim logErrors As String: logErrors = OUTPUT_PATH & STDERR
    
    If Not FileExists(logErrors) Then
        errs = "Run Python script failed."
        GoTo OUTPUT
    End If
    
    errs = ReadTextFile(logErrors)
    
OUTPUT:
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")

    If errs = "" Then
        res.Add KEY_STATUS, True
        res.Add KEY_VALUE, ReadTextFile(logOutput)
    Else:
        res.Add KEY_STATUS, False
        res.Add KEY_VALUE, errs
    End If
    
    Set RunPython = res
    
    Set oShell = Nothing ' clean object
    
    ' remove log file
    Call ClearOutput
    
End Function


Sub GetConfig()
    '''
    ' get configuration data from main.cfg
    '
    '''
    Dim cfg_path As String, str As String
    
    ' default value if no find
    PYTHON_PATH = ""
    OUTPUT_PATH = ThisWorkbook.Path & "\" & "outputs" & "\"
    STDOUT = "output.log"
    STDERR = "errors.log"

    cfg_path = ThisWorkbook.Path & "\" & CFG_FILE
    If Not FileExists(cfg_path) Then Exit Sub
        
    Open cfg_path For Input As #1
    Do While Not EOF(1)
        Line Input #1, str
        str = Trim(str)
        
        ' skip empty line o comment line
        If str = "" Or str Like "[#]*" Then
        
        ' python path
        ElseIf str = "[python]" Then
            Line Input #1, str
            str = Trim(str)
            If str Like "[#]*" Then
                PYTHON_PATH = ""
            ElseIf str Like ".\*" Then
                PYTHON_PATH = ThisWorkbook.Path & Right(str, Len(str) - 1)
            ElseIf str Like "\*" Then
                PYTHON_PATH = ThisWorkbook.Path & str
            Else
                PYTHON_PATH = str
            End If
        
        ' output path
        ElseIf str = "[output]" Then
            Line Input #1, str
            str = Trim(str)
            If Not str Like "[#]*" Then OUTPUT_PATH = str & "\"
            
        ' standard output/error
        ElseIf str = "[stdout]" Then
            Line Input #1, str
            str = Trim(str)
            If Not str Like "[#]*" Then STDOUT = str
        ElseIf str = "[stderr]" Then
            Line Input #1, str
            str = Trim(str)
            If Not str Like "[#]*" Then STDERR = str
        End If
    Loop
    Close #1
End Sub


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
    Dim logOutput As String: logOutput = OUTPUT_PATH & STDOUT
    Dim logErrors As String: logErrors = OUTPUT_PATH & STDERR
    If FileExists(logOutput) Then Kill logOutput
    If FileExists(logErrors) Then Kill logErrors
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
