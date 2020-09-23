<div align="center">

## Pie Graph Generator


</div>

### Description

To create a pie chart for reporting by adding a number of segments with the relevant data.
 
### More Info
 
There is a complete project demonstrating how to use the class Pie chart so even the beginner should be able to use this. All code it well commented.


<span>             |<span>
---                |---
**Submitted On**   |2002-04-10 21:46:14
**By**             |[Trevor Newsome](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/trevor-newsome.md)
**Level**          |Advanced
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Pie\_Graph\_708094102002\.zip](https://github.com/Planet-Source-Code/trevor-newsome-pie-graph-generator__1-33686/archive/master.zip)

### API Declarations

```
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
```





