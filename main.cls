VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim codemodules() As CModuleInfo
    Dim var As CVariableInfo
    Set var = New CVariableInfo
    
    Dim m As CModuleInfo
    Set m = New CModuleInfo
    
    m.Name = "hello"
    m.SegStartLine = 10
    m.SegEndLine = 100
    m.Variables.Add var, "key"
    m.Variables.Add var, "key"
    
End Sub
