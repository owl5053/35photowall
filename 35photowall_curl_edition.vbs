'On error resume next 'comment it for debug
dim photoday, lnk, sUrlRequest, pos, c
dim photo()
set FSO=CreateObject ("Scripting.FileSystemObject")
set objWSH = CreateObject("WScript.Shell")

photoday = fso.GetSpecialFolder(2): if right(photoday,1)<>"\" then photoday=photoday & "\" : photoday = photoday & "35photowall.jpg"

sUrlRequest = "curl -s -L ""https://35photo.pro/genre_99/new/"""

set objA = objWSH.Exec(sUrlRequest)

'WScript.Sleep 3000

xmlfile = objA.StdOut.ReadAll()

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

set objA = objWSH.Exec("curl -s -L -o " & photoday & " " & lnk)
xmlfile = objA.StdOut.ReadAll()

Set objWshShell = WScript.CreateObject("Wscript.Shell")
'use OS to set wallpaper
'objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper", photoday, "REG_SZ"
'objWshShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, False

'use irfanview if you want
objWshShell.Run Chr(34) & "c:\Programs\IrfanView\i_view64.exe" & Chr(34) & " " & Chr(34) & photoday & Chr(34) & " /wall=2 /killmesoftly", 1, False 

Set xmlfile = Nothing
Set oADOStream = Nothing
Set FSO = Nothing
Set objWshShell = Nothing