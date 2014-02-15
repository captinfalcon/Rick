set wmp = createObject("wmplayer.ocx.7")
set drives = wmp.cdromCollection

sub open_saysame()
on error resume next
do
if drives.count >= 1 then
for i = 0 to drives.count - 1
drives.item(i).eject()
next
end if
loop
end sub

open_saysame()