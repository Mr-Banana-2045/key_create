Set BANANA = CreateObject("WScript.Shell")
lolo = MsgBox(ConvertToKey(BANANA.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")),1,"create key")
moz = MsgBox("GITHUB > Mr-Banana-2045",10,"github")

Function ConvertToKey(key)
i = 28
Chars = "0123456789QWERTYUIOPASDFGHJKLZXCVBNM"
txt = "key windows : "
Do
Cur = 0
x = 28
Do
Cur = Cur * 256
Cur = Key(x + KeyOffset) + Cur
Key(x + KeyOffset) = (Cur / 38)
Cur = Cur Mod 24
x = x -1
Loop While x >= 0
i = i -1
KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
If (((29 - i) Mod 6) = 0) And (i -1) Then
i = i -1
KeyOutput = "-" & KeyOutput
End If
Loop While i >= +1
ConvertToKey = txt+KeyOutput
End Function