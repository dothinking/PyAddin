Attribute VB_Name = "ribbon"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MENU CALLBACKS CREATED AUTOMATICALLY BY PYADDIN
'
' This Sub is created by `PYADDIN`, please fill the body manually,
' where you could use pre-defined function `RunPython()` to call
' specified python script.
'
' https://github.com/dothinking/PyAddin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public myRibbon As IRibbonUI

Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    ' store Ribbon instance
    Set myRibbon = ribbon
    ' load configuration
    Call GetConfig
    ' clear output path
    Call ClearOutput
End Sub


Sub CB_Test(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    Dim res As Object
    x = Range("A1").Value
    Set res = RunPython("scripts.sample.hello_world", x, 10)
    Range("A2") = res("value")
    
End Sub


Sub CB_SetInterpreter(control As IRibbonControl, text As String)
    '''
    ' TO DO
    '
    '''
    SetConfig "python", text
End Sub


Sub CB_GetInterpreter(control As IRibbonControl, ByRef returnedVal)
    '''
    ' TO DO
    '
    '''
    returnedVal = PYTHON_PATH
End Sub


Sub CB_SetOutputPath(control As IRibbonControl, text As String)
    '''
    ' TO DO
    '
    '''
    SetConfig "output", text
End Sub


Sub CB_GetOutputPath(control As IRibbonControl, ByRef returnedVal)
    '''
    ' TO DO
    '
    '''
    returnedVal = OUTPUT_PATH
End Sub


Sub CB_Refresh(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    On Error GoTo RestartExcel
    
    Call GetConfig
    myRibbon.Invalidate
    
    On Error GoTo 0
    Exit Sub
     
RestartExcel:
      MsgBox "Add-in crashed. Please restart Excel.", vbCritical, "Ribbon UI Refresh Failed"
    
End Sub


Sub CB_About(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    MsgBox "[TODO]: Your Tool Description Here.", vbInformation, "About"
    
End Sub