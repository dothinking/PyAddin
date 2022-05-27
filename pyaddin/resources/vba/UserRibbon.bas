Attribute VB_Name = "UserRibbon"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MENU CALLBACKS CREATED AUTOMATICALLY BY PYADDIN
'
' Please add subroutines here to implement your Add-in features, where you could use 
' pre-defined function `RunPython()` to call specified python script. Two examples:
'
' Sub CB_Sample_1(control As IRibbonControl)
'     '''onAction for control: Sample 1'''
'     Dim res As Object
'     Dim x As Integer: x = Range("A1").Value
'     Dim y As Integer: y = Range("A2").Value'
'     Set res = RunPython("scripts.sample.run_example_1", x, y)
'     Range("A3") = res("value")'
' End Sub
'
' Sub CB_Sample_2(control As IRibbonControl)
'     '''onAction for control: Sample 2'''
'     RunPython "scripts.sample.run_example_2"
' End Sub
'
'
' https://github.com/dothinking/PyAddin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CB_Sample_1(control As IRibbonControl)
    '''onAction for control: Sample 1'''
    Dim res As Object
    Dim x As Integer: x = Range("A1").Value
    Dim y As Integer: y = Range("A2").Value
    
    Set res = RunPython("scripts.sample.run_example_1", x, y)
    Range("A3") = res("value")
    
End Sub


Sub CB_Sample_2(control As IRibbonControl)
    '''onAction for control: Sample 2'''
    RunPython "scripts.sample.run_example_2"
End Sub