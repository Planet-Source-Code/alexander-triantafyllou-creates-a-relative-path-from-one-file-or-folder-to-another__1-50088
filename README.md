<div align="center">

## Creates a relative path from one file or folder to another\.


</div>

### Description

Creates a relative path from one file or folder to another.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alexander Triantafyllou](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alexander-triantafyllou.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alexander-triantafyllou-creates-a-relative-path-from-one-file-or-folder-to-another__1-50088/archive/master.zip)





### Source Code

```
Private Declare Function PathRelativePathTo Lib "shlwapi.dll" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long
Private Const MAX_PATH As Long = 260
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
'-----------------------------------------------------------
' Creates a relative path from one file or folder to another.
'
' made by Alexander Triantafyllou alextriantf@yahoo.gr
'
' usage relative_path=get_relative_path_to(root_path,file_path)
' get_relative_path_to("d:\a\b\c\d","d:\a\b\index.html") will return
' "..\..\index.html"
' use FILE_ATTRIBUTE_DIRECTORY if the path is a directory
' or FILE_ATTRIBUTE_NORMAL if the path is a file
'----------------------------------------------------------
Public Function get_relative_path_to(ByVal parent_path As String, ByVal child_path As String) As String
Dim out_str As String
Dim par_str As String
Dim child_str As String
out_str = String(MAX_PATH, 0)
par_str = parent_path + String(100, 0)
child_str = child_path + String(100, 0)
PathRelativePathTo out_str, par_str, FILE_ATTRIBUTE_DIRECTORY, child_str, FILE_ATTRIBUTE_NORMAL
out_str = StripTerminator(out_str)
'MsgBox out_str
get_relative_path_to = out_str
End Function
'Remove all trailing Chr$(0)'s
Function StripTerminator(sInput As String) As String
 Dim ZeroPos As Long
 ZeroPos = InStr(1, sInput, Chr$(0))
 If ZeroPos > 0 Then
  StripTerminator = Left$(sInput, ZeroPos - 1)
 Else
  StripTerminator = sInput
 End If
End Function
```

