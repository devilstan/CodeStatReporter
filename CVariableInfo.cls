VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVariableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'變數類別

Public Name As String
Public DesignChangeSymbols As Collection
Public RelatedModules As Collection
Public Action As Integer '動作類別=讀/寫 (尚未實現)


Private Sub Class_Initialize()
        Set DesignChangeSymbols = New Collection
        Name = ""
        Action = -1
End Sub

Private Sub Class_Terminate()
        Set DesignChangeSymbols = Nothing
End Sub

