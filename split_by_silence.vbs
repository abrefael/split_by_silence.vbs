

Dim oExec, strScript, wShell, sFileSelected, OUT, timeList, strFromProc
Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>"&_
	"FILE.click();new ActiveXObject('Scripting.FileSystemObject')."&_
	"GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
silenceDetectVars = "-55dB:d=3.5"
strScript = "cmd /c ""ffmpeg  -v warning -i """ &_
	sFileSelected & """ -af silencedetect=" & silenceDetectVars &_
	",ametadata=mode=print:file=-:key=lavfi.silence_start -vn -sn  "&_
	"-f s16le  -y NUL | findstr ""lavfi.silence_start="""""

Set oExec = wShell.Exec(strScript)
strFromProc = ""
Do
    strFromProc = strFromProc & "," & oExec.StdOut.ReadLine()
Loop While Not oExec.Stdout.atEndOfStream
timeList = replace(strFromProc,"lavfi.silence_start=","")
timeList = mid(timeList, 2)
OUT = Left(sFileSelected, InStrRev(sFileSelected,".") - 1) & "%03d.mp3"
strScript = "ffmpeg -v warning -i """ & sFileSelected & """ -c copy -map 0+&_
	"-f segment -segment_times """ & timeList & """ """ & OUT & """"

Set oExec = wShell.Exec(strScript)
