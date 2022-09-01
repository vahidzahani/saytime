Dim speaks, speech
hour_now=hour(time)
minute_now=minute(time)
speaks =  hour_now & " " & minute_now
Set speech=CreateObject("sapi.spvoice")
speech.Speak speaks
