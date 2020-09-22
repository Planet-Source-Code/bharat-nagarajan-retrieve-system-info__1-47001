<div align="center">

## Retrieve System Info


</div>

### Description

Retrieves System Information in a single line of code
 
### More Info
 
To retrieve system info we use all kinda code but there is a function in VB 6.0 called Environ() which takes in a single parameter which makes our lives bit easy. Like if you want to know the users Temporary Directory or System Directory via VB 6.0

System Information


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bharat Nagarajan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bharat-nagarajan.md)
**Level**          |Advanced
**User Rating**    |3.0 (12 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bharat-nagarajan-retrieve-system-info__1-47001/archive/master.zip)





### Source Code

```
Private Sub cmdSysInfo_Click()
'Will popup a message box displaying your computer name
 MsgBox(environ("computername"))
'Like that we can get it for the variables like
    'ALLUSERSPROFILE()
    'APPDATA()
    'CommonProgramFiles()
    'COMPUTERNAME()
    'ComSpec()
    'HOMEDRIVE()
    'HOMEPATH()
    'LOGONSERVER()
    'NUMBER_OF_PROCESSORS()
    'OS()
    'Os2LibPath()
    'Path()
    'PATHEXT()
    'PROCESSOR_ARCHITECTURE()
    'PROCESSOR_IDENTIFIER()
    'PROCESSOR_LEVEL()
    'PROCESSOR_REVISION()
    'ProgramFiles()
    'SystemDrive()
    'SystemRoot()
    'TEMP()
    'TMP()
    'USERDOMAIN()
    'USERNAME()
    'USERPROFILE()
    'windir()
End Sub
```

