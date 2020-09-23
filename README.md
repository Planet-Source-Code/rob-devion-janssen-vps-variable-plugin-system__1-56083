<div align="center">

## VPS \- Variable Plugin System


</div>

### Description

VPS is a Plugin system; You write ActiveX plugins and VPS will execute them by a few standard functions you must have in your project.

Currently it has 2 plugins in it; A MUD server which isn't finished and a Backup system which is reasonably finished (delete the ini to get the config screen when starting VPS)

All source is included for each plugin and VPS itself, VPS is a console based application and you should use CAP to make it able to run inside a cmd.exe instead of it's own console window.
 
### More Info
 
No input is necessary.

A bit about the plugins (you can find this text also in the plugins dir as a .txt file)

Plugin required functions...

DLL Name = VPS

Class name = Filename of DLL

fstrGetName (no args) return result should be a string giving the friendly name of the plugin

fstrGetFunction (no args) return result should be a string giving the function that initializes the plugin

fClose (no args) allows plugin to close itself before VPS terminates

registerparent (refObject) allows the plugin to be able to return calls.

Parent functions

Log(String) - logs to the log file and screen.

Logfile(String) - logs only to file

Logscreen(String) - logs only to screen

That should get you started :)

No output except screen and what the plugin does.

None so far.


<span>             |<span>
---                |---
**Submitted On**   |2004-08-29 02:21:22
**By**             |[Rob 'Devion' Janssen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rob-devion-janssen.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[VPS\_\-\_Vari179129992004\.zip](https://github.com/Planet-Source-Code/rob-devion-janssen-vps-variable-plugin-system__1-56083/archive/master.zip)








