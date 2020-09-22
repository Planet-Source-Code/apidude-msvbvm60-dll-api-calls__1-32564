<div align="center">

## MSVBVM60\.DLL API Calls


</div>

### Description

Much like kernel32.dll or user32.dll, MSVBVM60.DLL (or MSVBVM50 for VB5) contains a myriad of useful API function calls that you can utilise - although many of these are undocumented.

Just out of interest, I'd thought I'd post a few up here. Seeing as we all have to distribute that annoyingly chunky DLL with our apps, maybe we could start researching and put the thing to really good use!
 
### More Info
 
Usually two or more Variants

A Variant


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[apidude](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/apidude.md)
**Level**          |Intermediate
**User Rating**    |4.1 (49 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/apidude-msvbvm60-dll-api-calls__1-32564/archive/master.zip)

### API Declarations

```
''add two variants together
Declare Function vbaVarAdd Lib "MSVBVM60.DLL" Alias "__vbaVarAdd" (var1 As Variant, var2 As Variant) As Variant
''subtract two variants
Declare Function vbaVarSub Lib "MSVBVM60.DLL" Alias "__vbaVarSub" (var1 As Variant, var2 As Variant) As Variant
''multiply two variants
Declare Function vbaVarMul Lib "MSVBVM60.DLL" Alias "__vbaVarMul" (var1 As Variant, var2 As Variant) As Variant
''concatenate two varients
Declare Function vbaVarCat Lib "MSVBVM60.DLL" Alias "__vbaVarCat" (var1 As Variant, var2 As Variant) As Variant
```


### Source Code

```
''you'll need to stick a command button on your form
Private Sub Command1_Click()
  ''See what these functions return...
  MsgBox "1+2=" & Trim$(Str$(vbaVarAdd(2, 1)))
  MsgBox "10-5=" & Trim$(Str$(vbaVarSub(5, 10)))
  MsgBox "6*6=" & Trim$(Str$(vbaVarMul(6, 6)))
  MsgBox CStr(vbaVarCat("lo!", "Hel"))
End Sub
```

