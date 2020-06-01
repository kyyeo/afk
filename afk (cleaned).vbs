Option Explicit
dim x, y, z, ie

'Create Internet Explorer object
Set ie = CreateObject("InternetExplorer.Application")
ie.Offline = True

'Navigate to URL
ie.Navigate "about:blank"
Do While ie.Busy : WScript.Sleep : Loop
ie.document.Title = "Exit AFK loop?"
ie.document.body.innerHTML = "<button name ='exit' onClick=document.all('continue').value='no'>Exit</button>" _
	& "<button name='pause' onClick=document.all('pause').value='yes'>Pause</button>" _
	& "<button name='unpause' onClick=document.all('pause').value='no'>UnPause</button>" _
	& "<input name='continue' type='hidden' value='yes'>" _
	& "<div style='width:100%; background-color:#E2E2FC; padding-left:5px'><p>Start Date/Time: <span id='datetime'></span></p></div>" _
	& "<div style='width:100%; background-color:#FFFFFF; padding-left:5px'><p>Time Elapsed: <span id='elapse'></span></p></div>" _
	& "<div style='width:100%; background-color:#FFFFFF; padding-left:5px'><p>Status: <span id='PauseStatus'></span></p></div>"

ie.Height	= 300
ie.Width	= 500
ie.MenuBar	= False
ie.StatusBar	= False
ie.AddressBar	= False
ie.ToolBar	= False
ie.Visible	= True


Set x = CreateObject("WScript.Shell")

'set time
z = Now
ie.document.getElementById("datetime").innerHTML = z

'loop process
Do Until ie.document.all("continue").Value = "no"
	if ie.document.all("pause").Value = "yes" then
		Do Until ie.document.all("pause").Value = "no"
			ie.document.getElementById("PauseStatus").innerHTML = "On Pause!"
		Loop
	else
		ie.document.getElementById("PauseStatus").innerHTML = "Script is Active"
		
		'keep screen alive
		x.sendkeys "{PRTSC}"

		'calculate time elapsed
		y = Now
		ie.document.getElementById("elapse").innerHTML = Mid(DateDiff("s",z,y),1,10)

		'wait 5 seconds
		wscript.sleep 5000
	end if
Loop

ie.quit

'Unload objects
Set ie = nothing
Set x = nothing
