X=MsgBox("Your computer iz destroid :P",0+16,"HAVE FUN ;)")
Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Users\Doom AF\Downloads\viruses by tbfdile._\Windows 10 Sparta Remix.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close