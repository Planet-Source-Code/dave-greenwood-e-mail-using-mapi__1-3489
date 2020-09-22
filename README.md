<div align="center">

## E\-mail using MAPI


</div>

### Description

Sends an e-mail using MAPI
 
### More Info
 
place two mapi controls on a form

set mapi user name and password


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Greenwood](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-greenwood.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-greenwood-e-mail-using-mapi__1-3489/archive/master.zip)





### Source Code

```
MAPIsession1.SignOn
if mapisession1.sessionID <> 0 then
with mapimessages1
.sessionid = MapiSession1.sessionID
.compose
.recipdisplayname "YOUR NAME"
recipaddress = "me@myipdomain.com"
.msgsubject = "SUBJECT"
.msgnotetext= "Message"
.send false
end with
mapisession1.signoff
end if
```

