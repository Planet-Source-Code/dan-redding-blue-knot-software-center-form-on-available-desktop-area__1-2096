<div align="center">

## Center Form on \*available\* desktop area


</div>

### Description

Centers a form in the part of your desktop not taken up by the taskbar or other system toolbars. If toolbars take up half the screen, no problem. If they are on the sides or the top, no problem.
 
### More Info
 
Can be used for any form in your project.

If your form is too tall or wide for the available space, the form is positioned along the top/left edges so that at least the control menu can be accessed. You may want to modify the code so that the form is resized in this case.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Redding \- Blue Knot Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-redding-blue-knot-software.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-redding-blue-knot-software-center-form-on-available-desktop-area__1-2096/archive/master.zip)

### API Declarations

```
'Paste these declarations and procedure in a module
'Note the change in the declaration from what the API Viewer
'pasted in: ByVal has been removed from lpvParam to allow
'passing a RECT (User-defined type)
'Be careful of this if you use the API call for anything
'else in your program.
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
 (ByVal uAction As Long, _
 ByVal uParam As Long, _
 lpvParam As Any, _
 ByVal fuWinIni As Long) As Long
Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
Public Const SPI_GETWORKAREA = 48
Public Function CenterInWorkArea(frm As Form)
Dim lNewTop As Long, lNewLeft As Long
Dim WA As RECT, lReturn As Long
 'Get the work area in a RECTangle structure from the '
 'SystemParametersInfo API Call
 lReturn = SystemParametersInfo(SPI_GETWORKAREA, 0&, WA, 0&)
 'Convert the virtual coordinates to scale coordinates
 WA.Left = WA.Left * Screen.TwipsPerPixelX
 WA.Right = WA.Right * Screen.TwipsPerPixelX
 WA.Top = WA.Top * Screen.TwipsPerPixelY
 WA.Bottom = WA.Bottom * Screen.TwipsPerPixelY
 'WA.Bottom-WA.Top = Work Area Height
 lNewTop = ((WA.Bottom - WA.Top - frm.Height) / 2) + WA.Top
 'Top is off screen or hidden because form is taller than workspace; adjust
 If lNewTop < WA.Top Then lNewTop = WA.Top
 'WA.Right - WA.Left = Work Area Width
 lNewLeft = ((WA.Right - WA.Left - frm.Width) / 2) + WA.Left
 'Left is off screen or hidden because form is too wide for workspace; adjust
 If lNewLeft < WA.Left Then lNewLeft = WA.Left
 'Perfect Centering!
 frm.Move lNewLeft, lNewTop
End Function
```


### Source Code

```
'Place this in Form_Load() or wherever else you think it is appropriate ;)
 CenterInWorkArea Me
```

