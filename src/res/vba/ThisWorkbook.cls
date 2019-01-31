'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WORKBOOK EVENTS CREATED AUTOMATICALLY BY PYADDIN
'
' Get configuration and set temp directory when load this addin file,
' and clean the temp file when unload this addin
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
    
    ' check TEMP path
    TEMP_PATH = ThisWorkbook.Path & "\temp\"
    If Dir(TEMP_PATH, vbDirectory) = "" Then MkDir TEMP_PATH
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ''
    ' clean temp files
    '
    '''
    
    If TEMP_PATH <> "" And Dir(TEMP_PATH & "*", 16) <> "." Then Kill TEMP_PATH & "*"
    
End Sub
