
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ddns(dyndns v2) update vbscript v1.1 by Mr.Blinky                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This script will check your wan ip and update your ddns account accordingly.
' A log file is created to keep track of changes (same directory as the script)

' Add a daily task to windows task scheduler to execute this script on a daily
' basis and set advanced options to repeat every 1 hour or 15 minutes for a
' period of 24 hours. After setting up the task, execute it manually and check
' the log file to see if your ip was updated succesfully

' Don't forget to fill in your domain details here :

protocol = "https"              ' http or https
server   = "dyndns.strato.com"  ' your domain registrar's server
hostname = "mydomain.com"       ' the domain you've registered
username = "mydomain.com"       ' your username
password = "password"           ' your password

' list of sites that can be queried for your wan ip address :

dim sites(2)
sites(0)="http://www.echoip.com"
sites(1)="http://ifconfig.me/ip"
sites(2)="http://checkip.dyndns.com"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

logfile =left(wscript.scriptfullname,len(wscript.scriptfullname)-3) & "log"

oldip = getoldip
newip = getwanip
if oldip <> newip then
 if newip <> "" then
  log("New IP address : " & newip & " Server update response : " & updateip(newip))
 else
  log("Failed to retrieve your IP address. Last known IP address : " & oldip)
 end if
else
 'remove the "'" at the beginning of the line below to log no change events too
 'log("No IP address change, current IP address : " & oldip)
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function stripip (s)

 b = "(25[0-5]|2[0-4]\d|[01]?\d\d?)" 'byte mask
 d = "\."
 set re = createobject("VBScript.RegExp")
 re.global = false
 re.pattern = b & d & b & d & b & d & b
 set ip = re.Execute(s)
 if ip.count > 0 then stripip=ip(0).value
end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function getoldip

 set o = createobject("scripting.filesystemobject")
 set f = o.opentextfile(logfile,1,True)
 do until f.atendofstream
  s = t
  t  = f.readline
  if t <>"" then s = t
 loop
 getoldip = stripip(s)
 f.close
end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function getwanip

 set o = createobject("MSXML2.XMLHTTP")
 for each site in sites
  o.open "GET", site, false
  o.send
  getwanip = stripip(o.responsetext)
  if getwanip <> "" then exit for
 next
end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function updateip(ip)

 url = protocol & "://" & username & ":" & password & "@" & server & "/nic/update?hostname=" & hostname & "&myip=" & ip & "&wildcard=nochg&mx=nochg&backmx=nochg"
 set o = createobject("MSXML2.XMLHTTP")
  o.open "GET", url, false
  o.send
  updateip = o.responsetext
end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function Log(s)

 set o = createobject("scripting.filesystemobject")
 set f = o.opentextfile(logfile,8,True)
 f.writeline(now & " " & s)
 f.close
end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
