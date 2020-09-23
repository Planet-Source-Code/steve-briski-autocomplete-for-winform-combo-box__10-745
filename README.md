<div align="center">

## Autocomplete for winform combo box


</div>

### Description

Autocomplete for winforms combo boxes. This was found on MSDN and enhanced to prevent the user from entering a selection that isn't in the combo box list. The code finds and selects the first match for the text that the user entered--if no match is found it selects the first entry that is "next" after the entered text. The delete and backspace key are also handled.
 
### More Info
 
the name of the combo box and the event arg for the combo box

Add the module to your application, then from your form just call the AutoComplete function passing in the name of your combobox and the event arg.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Briski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-briski.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB\.NET
**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__10-3.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-briski-autocomplete-for-winform-combo-box__10-745/archive/master.zip)





### Source Code

```
Module Module1
  Public Sub AutoComplete(ByVal cbo As ComboBox, ByVal e As System.Windows.Forms.KeyEventArgs)
    'call this from your form passing in the name of your combobox and the event arg:
    '  AutoComplete(cboState, e)
    Dim iIndex As Integer
    Dim sActual As String
    Dim sFound As String
    Dim bMatchFound As Boolean
    'if backspace then remove the last character that was typed in and try to find a match.
    'note that the selected text from the last character typed in to the end of the combo text field will also be deleted.
    If e.KeyCode = Keys.Back Then
      cbo.Text = Mid(cbo.Text, 1, Len(cbo.Text) - 1)
    End If
    ' Do nothing for some keys such as navigation keys.
    If ((e.KeyCode = Keys.Left) Or _
      (e.KeyCode = Keys.Right) Or _
      (e.KeyCode = Keys.Up) Or _
      (e.KeyCode = Keys.Down) Or _
      (e.KeyCode = Keys.PageUp) Or _
      (e.KeyCode = Keys.PageDown) Or _
      (e.KeyCode = Keys.Home) Or _
      (e.KeyCode = Keys.End)) Then
      Return
    End If
    Do
      ' Store the actual text that has been typed.
      sActual = cbo.Text
      ' Find the first match for the typed value.
      iIndex = cbo.FindString(sActual)
      ' Get the text of the first match.
      'if index > -1 then a match was found.
      If (iIndex > -1) Then
        sFound = cbo.Items(iIndex).ToString()
        ' Select this item from the list.
        cbo.SelectedIndex = iIndex
        ' Select the portion of the text that was automatically
        ' added so that additional typing will replace it.
        cbo.SelectionStart = sActual.Length
        cbo.SelectionLength = sFound.Length
        bMatchFound = True
      Else
        'if there isn't a match and the text typed in is only 1 character or nothing then just select the first
        'entry in the combo box.
        If sActual.Length = 1 Or sActual.Length = 0 Then
          cbo.SelectedIndex = 0
          cbo.SelectionStart = 0
          cbo.SelectionLength = Len(cbo.Text)
          bMatchFound = True
        Else
          'if there isn't a match for the text typed in then remove the last character of the text typed in
          'and try to find a match.
          cbo.SelectionStart = sActual.Length - 1
          cbo.SelectionLength = sActual.Length - 1
          cbo.Text = Mid(cbo.Text, 1, Len(cbo.Text) - 1)
        End If
      End If
    Loop Until bMatchFound
  End Sub
End Module
```

