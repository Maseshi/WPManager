set objshell = CreateObject("WScript.Shell")

path = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
digitalID = objshell.RegRead(path & "DigitalProductId") 
wscript.echo(convertToKey(digitalID))

function convertToKey(key) 
    const keyOffset = 52 
    dim isWin8, maps, i, j, current, keyOutput, last, keypart1, insert 
    
    'Check if OS is Windows 8 
    isWin8 = (key(66) \ 6) and 1 
    key(66) = (key(66) and &HF7) Or ((isWin8 and 2) * 4) 
    i = 24 
    maps = "BCDFGHJKMPQRTVWXY2346789" 
    do 
        current = 0 
        j = 14 
        do 
           current = current * 256 
           current = key(j + keyOffset) + current 
           key(j + keyOffset) = (current \ 24) 
           current = current mod 24 
            j = j -1 
        loop while j >= 0 
        i = i -1 
        keyOutput = mid(maps, current + 1, 1) & keyOutput 
        last = current 
    loop while i >= 0  
     
    if (isWin8 = 1) then 
        keypart1 = mid(keyOutput, 2, last) 
        insert = "N" 
        keyOutput = replace(keyOutput, keypart1, keypart1 & insert, 2, 1, 0) 
        if last = 0 then keyOutput = insert & keyOutput 
    end if
    convertToKey = mid(keyOutput, 1, 5) & "-" & mid(keyOutput, 6, 5) & "-" & mid(keyOutput, 11, 5) & "-" & mid(keyOutput, 16, 5) & "-" & mid(keyOutput, 21, 5) 
end function