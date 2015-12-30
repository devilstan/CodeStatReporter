Attribute VB_Name = "ModuleFoo"
Function LTrimEx(str)
'資料來源：http://stackoverflow.com/questions/1098606/trim-leading-spaces-including-tabs
'說明：This function removes all leading whitespace (spaces, tabs etc) from a string
    Dim re As String
    LTrimEx = Replace(LTrim(str), Chr(13), "")
End Function
