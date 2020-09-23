<div align="center">

## modBDimg


</div>

### Description

VB module to read and store ANY kind of picture file pictureboxes support into a database, it's easy and it's pretty fast... please, rate it, any comments are wellcome
 
### More Info
 
FileName: Name of the picture file to store

rsImg: recordset with a memo field

FieldName: Name of the memo field to use

SaveImage: Nothing

ReadImage: An IPictureDisp object assignable to a picturebox or image

uses temporary storage


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Manuel Fernandez](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/manuel-fernandez.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/manuel-fernandez-modbdimg__1-14470/archive/master.zip)

### API Declarations

```
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
```


### Source Code

```

'Saves the image Filename (any kind Picturebox supports: jpg, gif, ico, bmp, wmf..) in to
'the current record of the recordset rsImg, using the field FieldName (must be a memo field!!!)
'USE: SaveImage("c:\sample.gif", rs)
Public Sub SaveImage(Filename As String, rsImg As Recordset, Optional FieldName As String = "Image")
  On Error Goto EH
  Dim fh As Integer
  Dim strFile As String
  If rsImg.BOF Or rsImg.EOF Then Err.Raise vbObjectError + 1, "SaveImage", "EOF or BOF encountered"
  fh = FreeFile
  Open Filename For Binary Access Read As fh
  strFile = String(LOF(fh), " ")
  Get fh, , strFile
  Close fh
  rsImg(FieldName) = strFile
  Exit Sub
EH:
End Sub
'Reads the image (any kind Picturebox supports: jpg, gif, ico, bmp, wmf..) from
'the current record of the recordset rsImg, using the field FieldName, and returns it.
'USE: picture1.picture=ReadImage(rsImg)
Public Function ReadImage(rsImg As Recordset, Optional FieldName As String = "Image") As IPictureDisp
  On Error Goto EH
  Dim strFile As String
  Dim fh As Integer
  If rsImg.BOF Or rsImg.EOF Then Err.Raise vbObjectError + 2, "EeadImage", "EOF or BOF encountered"
  ChDir App.Path
  strFile = rsImg(FieldName)
  fh = FreeFile
  Open GetTempDir & "tmpimage.temp" For Binary Access Write As fh
  Put #fh, , strFile
  Close fh
  Set LeerImagen = LoadPicture(GetTempDir & "tmpimage.temp")
  Kill GetTempDir & "tmpimage.temp"
  Exit Function
EH:
End Function
Private Function GetTempDir() As String
  GetTempDir = String(255, " ")
  GetTempPath 255, GetTempDir
  GetTempDir = Left(Trim(GetTempDir), Len(Trim(GetTempDir)) - 1)
End Function
```

