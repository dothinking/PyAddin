VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
