<div align="center">

## Stop New Spam with ASP


</div>

### Description

This code takes a submitted email address/string and converts it to standardized ASCII Character codes, thus blocking bots/spiders from reading your email address, while yet allowing visitors to continue to read and use your proper address via mailto links, etc
 
### More Info
 
Input a String

Returns an ASCII Encoded String


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Teco](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/teco.md)
**Level**          |Beginner
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/teco-stop-new-spam-with-asp__4-7677/archive/master.zip)

### API Declarations

```
EXAMPLE USE
------------------------
&lt;A HREF="mailto:&lt;%= ASCIICode("a@b.com") %&gt;"&gt;&lt;%= ASCIICode("a@b.com") %&gt;&lt;/A&gt;
```


### Source Code

```
Function ASCIICode(strInput)
  strOutput = ""
  For i = 1 To Len(strInput)
    strOutput = strOutput & "&#" & ASC(Mid(strInput, i, 1)) & ";"
  Next
  ASCIICode = strOutPut
End Function
```

