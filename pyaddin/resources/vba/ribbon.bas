Attribute VB_Name = "Ribbon"
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
    SetConfig "ribbon", ObjPtr(ribbon)
    
    ' load configuration
    Call LoadConfig
    
    ' clear output path
    Call ClearOutput
End Sub


Sub CB_Test_1(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    Dim res As Object
    Dim x As Integer: x = Range("A1").Value
    Dim y As Integer: y = Range("A2").Value
    
    Set res = RunPython("scripts.sample.hello_world_1", x, y)
    Range("A3") = res("value")
    
End Sub



Sub CB_Test_2(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    RunPython "scripts.sample.hello_world_2"
End Sub


Sub CB_SetInterpreter(control As IRibbonControl, text As String)
    '''
    ' TO DO
    '
    '''
    SetConfig "python", text
    PYTHON_PATH = text
End Sub


Sub CB_GetInterpreter(control As IRibbonControl, ByRef returnedVal)
    '''
    ' TO DO
    '
    '''
    returnedVal = PYTHON_PATH
End Sub


Sub CB_Refresh(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    Call LoadConfig
    
    ' restore Ribbon instance in case VBA script was stop unexpectedly
    If myRibbon Is Nothing Then Set myRibbon = RestoreRibbon()
    myRibbon.Invalidate
    
End Sub


Sub CB_About(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    MsgBox "[TODO]: Your Tool Description Here.", vbInformation, "About"
    
End Sub