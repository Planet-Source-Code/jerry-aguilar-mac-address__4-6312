<div align="center">

## MAC address


</div>

### Description

Get the clients MAC(Media Access Control)

address, a hardware address that uniquely

identifies each node of a network. Works great on

LAN's. Firewalls and Proxy's will be an issue

depending what side of them you're coding for.
 
### More Info
 
You can't navigate to it running PWS on the same

pc but if you are running PWS, you can navigate

to it from another pc on the same lan (it does

not like 127.0.0.1)

Returns the client IP and MAC address.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jerry Aguilar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jerry-aguilar.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jerry-aguilar-mac-address__4-6312/archive/master.zip)

### API Declarations

```
This code is AS IS! I had a need for it on a
project I was working on and found almost no info
anywhere on what I needed to accomplish. If it
helps you, great! If it does not work for you:
1.Make sure you're not trying to hit on the same
pc it's on.
2.Comment out where the file gets deleted
(fso.deletefile "c:\" & strIP & ".txt"), to view
some potential error info.
3.Have fun debugging :) (I did)
```


### Source Code

```
<%@ LANGUAGE="VBSCRIPT"%>
<%
	strIP = Request.ServerVariables("REMOTE_ADDR")
	strMac = GetMACAddress(strIP)
	strHost = Request.ServerVariables("REMOTE_HOST")
 Function GetMACAddress(strIP)
 set net = Server.CreateObject("wscript.network")
 set sh = Server.CreateObject("wscript.shell")
 sh.run "%comspec% /c nbtstat -A " & strIP & " > c:\" & strIP & ".txt",0,true
 set sh = nothing
 set fso = createobject("scripting.filesystemobject")
 set ts = fso.opentextfile("c:\" & strIP & ".txt")
 macaddress = null
 do while not ts.AtEndOfStream
  data = ucase(trim(ts.readline))
  if instr(data,"MAC ADDRESS") then
  macaddress = trim(split(data,"=")(1))
  exit do
  end if
 loop
 ts.close
 set ts = nothing
 fso.deletefile "c:\" & strIP & ".txt"
 set fso = nothing
 GetMACAddress = macaddress
 End Function
%>
<html>
<HEAD>
<TITLE>Say Hello To the MAC MAN</TITLE>
</HEAD>
<body>
<%Response.Write("Your IP is : " & strIP & "<br>" & vbcrlf)%>
<%Response.Write("Your MAC is : " & strMac & vbcrlf)%>
</body>
</html>
```

