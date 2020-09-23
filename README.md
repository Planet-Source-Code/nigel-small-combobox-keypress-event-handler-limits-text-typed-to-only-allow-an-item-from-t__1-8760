<div align="center">

## ComboBox KeyPress Event Handler \- limits text typed to only allow an item from the list


</div>

### Description

This combo box event handler will limit the text typed to only items available from the list. This essentially makes the control into a dropdown list but allows the items to be selected by typing as well as clicking.
 
### More Info
 
This will work on any ComboBox. It assumes the control is called Combo1 but obviously, this can be changed.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nigel Small](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nigel-small.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nigel-small-combobox-keypress-event-handler-limits-text-typed-to-only-allow-an-item-from-t__1-8760/archive/master.zip)





### Source Code

```
' function to intercept keypresses to combo box and allow
' only valid keys (such as are in list)
'
Private Sub Combo1_KeyPress(KeyAscii As Integer)
 Dim NewText As String
 Dim ValidCount As Integer
 Dim ValidValue As String
 ' do only if key pressed is printable character
 If KeyAscii >= 32 And KeyAscii <> 127 Then
  ' predict new text after keypress
  NewText = LCase(Left(Combo1.Text, Combo1.SelStart) + Chr(KeyAscii) + Mid(Combo1.Text, Combo1.SelStart + Combo1.SelLength + 1))
  ' find number of matches in combo list
  ValidCount = 0
  ValidValue = ""
  For i = 0 To Combo1.ListCount - 1
   If NewText = LCase(Left(Combo1.List(i), Len(NewText))) Then
    ValidCount = ValidCount + 1
    ValidValue = Combo1.List(i)
   End If
  Next
  ' cancel keypress if invalid
  If ValidCount <= 1 Then KeyAscii = 0
  ' select if one match only
  If ValidCount = 1 Then
   Combo1.Text = ValidValue
   Combo1.SelStart = 0
   Combo1.SelLength = Len(ValidValue)
  End If
 End If
End Sub
```

