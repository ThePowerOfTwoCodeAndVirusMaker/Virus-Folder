Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Users\Doom AF\Downloads\viruses by tbfdile._\Windows 10 Sparta Remix(1).wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
X=MsgBox(“Message Description”,0+16,”Title”)