<div align="center">

## Alternate ShrinkFormUnload


</div>

### Description

Shrink Window on Unload
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim Markoski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-markoski.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-markoski-alternate-shrinkformunload__1-28214/archive/master.zip)





### Source Code

```
Public Sub SqueezeWindow_FormUnload(p_FrmCurrent As Form, p_dblIncrement As Double)
' Comments :
' Parameters: p_FrmCurrent
'       p_dblIncrement -
' Modified :
'
' -------------------------------
On Error GoTo PROC_ERR
Do While (p_FrmCurrent.Height > 405 Or p_FrmCurrent.Width > 1680)
p_FrmCurrent.Height = p_FrmCurrent.Height - p_dblIncrement
p_FrmCurrent.Width = p_FrmCurrent.Width - p_dblIncrement
DoEvents
Loop
DoEvents
PROC_EXIT:
p_FrmCurrent.Hide
DoEvents
Unload p_FrmCurrent
Exit Sub
PROC_ERR:
MsgBox Err.Description
Resume PROC_EXIT
End Sub
```

