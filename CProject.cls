VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'程式碼專案

Public ModuleSight As Collection
Public VariableSight As Collection

Private Sub Class_Initialize()
    Set ModuleSight = New Collection
    Set VariableSight = New Collection
End Sub

Private Sub Class_Terminate()
    Set ModuleSight = Nothing
    Set VariableSight = Nothing
End Sub

Public Sub PrintStat(sh As Worksheet)

End Sub

