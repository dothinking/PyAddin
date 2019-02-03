'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WORKBOOK EVENTS CREATED AUTOMATICALLY BY PYADDIN
'
' Get configuration when loading this addin file -> main.cfg
'
' https://github.com/dothinking/PyAddin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Workbook_Open()
    '''
    ' set global parameters when open this addin
    '
    '''
    
    ' check python path
    Call GetConfig
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ''
    ' clean output files
    '
    '''
    On Error Resume Next
    If OUTPUT_PATH <> "" And Dir(OUTPUT_PATH, vbDirectory) <> Empty Then Kill OUTPUT_PATH & "*.*"
    
End Sub
