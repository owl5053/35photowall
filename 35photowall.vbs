'On error resume next 'comment it for debug
dim photoday, lnk, sUrlRequest, pos, c
dim photo()
set FSO=CreateObject ("Scripting.FileSystemObject")
photoday = fso.GetSpecialFolder(2): if right(photoday,1)<>"\" then photoday=photoday & "\" : photoday = photoday & "35photowall.jpg"

sUrlRequest = "https://35photo.pro/genre_99/new"' also you can set 98 - only 18+ :)
Set oXMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
oXMLHTTP.Open "GET", sUrlRequest, False

oXMLHTTP.setRequestHeader "Cookie", "nude=true"
oXMLHTTP.Send
xmlfile=oXMLHTTP.Responsetext
Set oXMLHTTP = Nothing


pos=1
c=1

do
ef=instr(pos,lcase(xmlfile),"_800n.jpg")

if ef<>0 then 
	beg=instrrev(lcase(xmlfile),"https://c1.35photo.pro",ef)
	lnk=mid(xmlfile,beg,ef-beg+9)
	lnk=replace(lnk,"_temp","_main")
	lnk=replace(lnk,"/sizes","")
	lnk=replace(lnk,"_800n.jpg",".jpg")
	ReDim Preserve photo(c)
	photo(c)=lnk
	pos=ef+1
	c=c+1
end if
loop while ef<>0

randomize
lnk=photo(1+int(rnd*UBound(photo)))


Set oXMLHTTP2 = CreateObject("WinHttp.WinHttpRequest.5.1")
oXMLHTTP2.Open "GET", lnk, False
oXMLHTTP2.Send
Set oADOStream = CreateObject("ADODB.Stream")
oADOStream.Mode = 3 'RW
oADOStream.Type = 1 'Binary
oADOStream.Open
oADOStream.Write oXMLHTTP2.ResponseBody

oADOStream.SaveToFile photoday, 2

Set objWshShell = WScript.CreateObject("Wscript.Shell")
'use OS to set wallpaper
Set objFile = FSO.GetFile(photoday)
Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.Namespace(FSO.GetParentFolderName(objFile))
objFolder.ParseName(FSO.GetFileName(objFile)).InvokeVerb "setdesktopwallpaper"

'use irfanview if you want
'objWshShell.Run Chr(34) & "c:\Programs\IrfanView\i_view64.exe" & Chr(34) & " " & Chr(34) & photoday & Chr(34) & " /wall=3 /killmesoftly", 1, False 

Set oXMLHTTP2 = Nothing
Set oADOStream = Nothing
Set FSO = Nothing
Set objWshShell = Nothing
