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
'選擇目錄
    Dim diaFolder As FileDialog
    Dim SearchPath As String
    
    SearchPath = ""

    ' 打開檔案選擇視窗
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show

    '取得選擇路徑
    SearchPath = diaFolder.SelectedItems(1)
    
    '列出路徑下的檔案
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim c As Range
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(SearchPath)
    Set c = Range("M2")
    Range("M:N").ClearContents
    Range("M1") = "原始碼檔案"
    Range("N1") = "路徑"
    'loops through each file in the directory and prints their names and path
    For Each objFile In objFolder.Files
        'print file name
        c.Value = objFile.Name
        'print file path
        c.Offset(0, 1).Value = objFile.Path
        Set c = c.Offset(1, 0)
    Next objFile

    Set diaFolder = Nothing
End Sub

Private Sub CommandButton2_Click()
'執行分析
    Dim FileNum As Integer
    Dim DataLine As String
    Dim file_beg As Range, file_end As Range
    
    Set file_beg = Range("N2")
    Set file_end = Range("N1").End(xlDown)
    
    '逐一載入原始碼
    For Each f In Range(file_beg, file_end)
        Dim row_idx As Integer
        Dim line_trim As String
        f.Interior.ColorIndex = 6
        
        Application.ScreenUpdating = False
        FileNum = FreeFile()
        Open f.Value For Input As #FileNum
        '逐行讀取, 並載入至A行
        Range("A:A").ClearContents
        Range("A:A").Font.ColorIndex = xlAutomatic
        row_idx = 1
        While Not EOF(FileNum)
            Line Input #FileNum, DataLine ' read in data 1 line at a time
            ' decide what to do with dataline,
            ' depending on what processing you need to do for each case
            line_trim = LTrimEx(DataLine)
            If Left(line_trim, 1) = ";" Then
                Cells(row_idx, 1).Font.ColorIndex = 10
            End If
            Cells(row_idx, 1).Value = DataLine
            row_idx = row_idx + 1
        Wend
        Close #FileNum
        
        Application.ScreenUpdating = True
        f.Interior.ColorIndex = 2
    Next
    
    
End Sub

