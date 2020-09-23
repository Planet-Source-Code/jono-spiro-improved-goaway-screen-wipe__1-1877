<div align="center">

## \*Improved\* GoAway Screen Wipe


</div>

### Description

I just quickly improved this screen wipe function written by Jesse Foster.

It runs a bit smoother and is usable as a function for any form. Just drop it into a module. Also, I made it work so that it is 'public' not private.
 
### More Info
 
'Screenwipe form to wipe, speed

'example: screenwipe form1,100


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jono Spiro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jono-spiro.md)
**Level**          |Unknown
**User Rating**    |4.2 (159 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jono-spiro-improved-goaway-screen-wipe__1-1877/archive/master.zip)





### Source Code

```
Public Function screenWipe(Form As Form, CutSpeed As Integer) As Boolean
 Dim OldWidth As Integer
 Dim OldHeight As Integer
 Form.WindowState = 0
 If CutSpeed <= 0 Then
 MsgBox "You cannot use 0 as a speed value"
 Exit Function
 End If
 Do
 OldWidth = Form.Width
 Form.Width = Form.Width - CutSpeed
 DoEvents
 If Form.Width <> OldWidth Then
  Form.Left = Form.Left + CutSpeed / 2
  DoEvents
 End If
 OldHeight = Form.Height
 Form.Height = Form.Height - CutSpeed
 DoEvents
 If Form.Height <> OldHeight Then
  Form.Top = Form.Top + CutSpeed / 2
  DoEvents
 End If
 Loop While Form.Width <> OldWidth Or Form.Height <> OldHeight
 Unload Form
End Function
```

