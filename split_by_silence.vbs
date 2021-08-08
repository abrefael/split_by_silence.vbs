'######################################################################################
'VBScript script to split mp3 file according to silence bits at selected lengths. This
'is based on the work of Vitaly Shukela, who created an shell script for unix based
'systems. You can find his wrk here:
'	https://gist.github.com/vi/2fe3eb63383fcfdad7483ac7c97e9deb
'This was done to to the lack of unix\linux based systems in the business world in my
'country.
'NOTE: you should run this script from a folder containing ffmpeg binaries or you
'should have your ffmpeg directory in your windows "PATH" variable.
'######################################################################################
' Some variables neede for this to work
Dim oExec, strScript, wShell, sFileSelected, OUT, timeList, strFromProc
Set wShell=CreateObject("WScript.Shell")

' This allows you to select your MP3 file (a file explorer shell)
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>"&_
	"FILE.click();new ActiveXObject('Scripting.FileSystemObject')."&_
	"GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine

' Here you can change your silencedetect preferences (i.e. the volume level of "silnce"
' for me, after increasing the sound of the file it was somewhere around 50.1 dB, so I
' set it to 50dB and the length of silnce was approximately 4 seconds, so I entered 3.5
' seconds. You can change this to your prefences.

silenceDetectVars = "-48dB:d=3.5"

'First thing is to recognise the "silences" in the file
strScript = "cmd /c ""ffmpeg  -v warning -i """ &_
	sFileSelected & """ -af silencedetect=" & silenceDetectVars &_
	",ametadata=mode=print:file=-:key=lavfi.silence_start -vn -sn  "&_
	"-f s16le  -y NUL | findstr ""lavfi.silence_start="""""

Set oExec = wShell.Exec(strScript)
strFromProc = ""
Do
	If strFromProc = "" Then
		strFromProc = oExec.StdOut.ReadLine()
	Else
		strFromProc = strFromProc & "," & oExec.StdOut.ReadLine()
	End If
Loop While Not oExec.Stdout.atEndOfStream

timeList = replace(strFromProc,"lavfi.silence_start=","")
'timeList = mid(timeList, 2)

'Just letting you know...
msgbox ("Splitting points are: " & timeList)

'Creating a template for the name of the files after the split (You cna actually
'change this if you wish...)
'The format is: Name_of_input_filexxx.mp3, where xxx is an ascending index from
'000 onward. In order to change this, I recomend not to touch the "%03d.mp3" part

OUT = Left(sFileSelected, InStrRev(sFileSelected,".") - 1) & "%03d.mp3"
'Running the split

strScript = "ffmpeg -v warning -i """ & sFileSelected & """ -c copy -map 0 "&_
	"-f segment -segment_times """ & timeList & """ """ & OUT & """"

Set oExec = wShell.Exec(strScript)
