Attribute VB_Name = "general"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GENERAL FUNCTION FOR PYADDIN
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' global parameters
Public Const CFG_FILE = "main.cfg"

Public TEMP_PATH As String 'defined on the loading of this addin
Public PYTHON_PATH As String

Function RunPython(method_name As String, args, ByRef res As String) As Boolean
    '''
    ' run python script
    ' :param method_name: a string refer to the called python method -> package.module.method
    ' :param args: array for python arguments
    ' :param res: python return string
    ' :returns: True if everything is OK else False
    '
    '''
    Dim oShell As Object, cmd As String, str_args As String
    Dim log_output As String, log_errors As String, errs As String
    
    ' check python path
    If PYTHON_PATH = "" Then
        errs = "Please set Python path first: " & ThisWorkbook.Path & "\" & CFG_FILE
        GoTo OUTPUT
    ElseIf Not PYTHON_PATH Like "*.exe" Then
        PYTHON_PATH = PYTHON_PATH & ".exe"
        If Dir(PYTHON_PATH, 16) = Empty Then
            errs = "Could not find Python: " & PYTHON_PATH
            GoTo OUTPUT
        End If
    End If
    
    ' join command
    PYTHON = """" & PYTHON_PATH & """ "
    main = """" & ThisWorkbook.Path & "\main.py"" "
    For i = LBound(args) To UBound(args)
        str_args = str_args & " """ & args(i) & """"
    Next
    cmd = PYTHON & main & method_name & str_args
    
    ' execute command
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run cmd, 0, 1
    
    ' results
    log_output = TEMP_PATH & "output.log"
    log_errors = TEMP_PATH & "errors.log"
    errs = ReadTextFile(log_errors)
    
OUTPUT:

    If errs = "" Then
        RunPython = True
        res = ReadTextFile(log_output)
    Else:
        RunPython = False
        res = errs
    End If
    
    Set oShell = Nothing ' clean object
    
    ' remove log file
    If log_errors <> "" And Dir(log_errors, 16) <> Empty Then
        Kill log_output
        Kill log_errors
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

Sub GetConfig()
    '''
    ' get configuration data from main.cfg
    '
    '''
    Dim str As String
        
    Open ThisWorkbook.Path & "\" & CFG_FILE For Input As #1
    Do While Not EOF(1)
        Line Input #1, str
        str = Trim(str)
        If str = "" Or str Like "[#]*" Then
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
        End If
    Loop
    Close #1
End Sub
