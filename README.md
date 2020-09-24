<div align="center">

## Access Linked Table Path Correction


</div>

### Description

When the folder of an access application and its linked table source changes, this code will modify the linked tables to comply.
 
### More Info
 
In order that code to work the Access application and the linked table source file must be in the same directory. Also the user running the application must have the permission to modify the tables.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Yalin Meric](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/yalin-meric.md)
**Level**          |Advanced
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VBA MS Access
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/yalin-meric-access-linked-table-path-correction__1-48743/archive/master.zip)





### Source Code

```
Sub TabloLinkleriniKontrolEt(sSourceFile As String)
 Dim daTaban As Database, tbTablo As TableDef
 Set daTaban = CurrentDb
 For Each tbTablo In daTaban.TableDefs
 If InStr(tbTablo.Connect, "DATABASE=") > 0 Then
  Debug.Print tbTablo.Connect
  If tbTablo.Connect <> ";DATABASE=" & Application.CurrentProject.Path & "\" & sSourceFile Then
  tbTablo.Connect = ";DATABASE=" & Application.CurrentProject.Path & "\" & sSourceFile
  tbTablo.RefreshLink
  End If
 End If
 Next
 daTaban.Close
End Sub
```

