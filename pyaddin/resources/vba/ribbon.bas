Attribute VB_Name = "Ribbon"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MENU CALLBACKS CREATED AUTOMATICALLY BY PYADDIN
'
' This module is created by `PYADDIN`, supporting the default features, e.g. set Python
' interpreter, load/refresh configuration.
'
' https://github.com/dothinking/PyAddin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public myRibbon As IRibbonUI


Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    '''Initialization when loading Ribbon UI.'''
    ' store Ribbon instance
    Set myRibbon = ribbon
    SetConfig KEY_RIBBON, ObjPtr(ribbon)
    
    ' load configuration
    Call LoadConfig
    
    ' clear output path
    Call ClearOutput
End Sub


Sub CB_SetInterpreter(control As IRibbonControl, text As String)
    '''onChange for control: Python Interpreter'''
    SetConfig "python", text
    PYTHON_PATH = text
End Sub


Sub CB_GetInterpreter(control As IRibbonControl, ByRef returnedVal)
    '''getText for control: Python Interpreter'''
    returnedVal = PYTHON_PATH
End Sub


Sub CB_Refresh(control As IRibbonControl)
    '''onAction for control: Refresh Settings'''
    Call LoadConfig
    
    ' restore Ribbon instance in case VBA script was stop unexpectedly
    If myRibbon Is Nothing Then Set myRibbon = RestoreRibbon()
    myRibbon.Invalidate
    
End Sub


Sub CB_About(control As IRibbonControl)
    '''onAction for control: About'''
    MsgBox "[TODO]: Your Tool Description Here.", vbInformation, "About"    
End Sub