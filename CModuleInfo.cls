VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'副程式類別

Public Name As String
Public Code As String
Public FullPath As String
Public SegStartLine As Integer
Public SegEndLine As Integer
Public Additions As Integer
Public Comments As Integer

Public Variables As Collection

Private Sub Class_Initialize()
        Set Variables = New Collection
        Name = 0
        Code = 0
        FullPath = 0
        SegEndLine = 0
        SegEndLine = 0
        Additions = 0
        Comments = 0
End Sub

Private Sub Class_Terminate()
        Set Variables = Nothing
End Sub
