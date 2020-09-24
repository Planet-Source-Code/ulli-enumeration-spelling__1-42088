<div align="center">

## Enumeration Spelling


</div>

### Description

How to keep the upper/lower case of Enumeration members
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Beginner
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ulli-enumeration-spelling__1-42088/archive/master.zip)





### Source Code

```
Option Explicit
'Due to a bug in VB members of an Enumeration do
'not keep their upper/lower case when they are
'typed differently in code. There is a simple
'solution however. All you have to do is repeat
'the'member names with their proper spelling and
'case bracketed by
#If False Then
#End If
'as the following example shows:
Public Enum TTStyle
  TTStandardIfActive = TTS_STANDARD
  TTBalloonIfActive = TS_BALLOON
  TTStandardAlways = TTS_STANDARD Or TTS_ALWAYSTIP
  TTBalloonAlways = TTS_BALLOON Or TTS_ALWAYSTIP
  TTNone = -1
End Enum
#If False Then
Private TTStandardIfActive, TTBalloonIfActive, TTStandardAlways, TTBalloonAlways, TTNone 'to preserve the case
#End If
```

