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
    If OUTPUT_PATH <> "" And Dir(OUTPUT_PATH, vbDirectory) <> Empty Then Kill OUTPUT_PATH & "*.*"    
End Sub


Sub CB_Test(control As IRibbonControl)
    '''
    ' TO DO
    '
    '''
    Dim res$    
    Call RunPython("scripts.sample.hello_world", Array(), res)    
    Range("A1") = res
    
End Sub


Sub CB_SetInterpreter(control As IRibbonControl, text As String)
    '''
    ' TO DO
    '
    '''
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

