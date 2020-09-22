<div align="center">

## A1 \- 'Animated'  spinning globe using only a label and a timer\. \(WebDings Required\)


</div>

### Description

Using webdings, you can created a VERY rudamentary spinning globe. Please do not vote.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\[\]\)utch\[\]v\[\]aster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/utch-v-aster.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/utch-v-aster-a1-animated-spinning-globe-using-only-a-label-and-a-timer-webdings-required__1-54893/archive/master.zip)





### Source Code

```
Const G1 As String = "ü"
Const G2 As String = "ý"
Const G3 As String = "þ"
'Label needs to be WEBDINGS font, at a relatively large size.
'My Timer1.Interval is set to 150
Private Sub Timer1_Timer()
 If Label1 = G1 Then
  Label1 = G2
 ElseIf Label1 = G2 Then
  Label1 = G3
 Else
  Label1 = G1
 End If
End Sub
```

