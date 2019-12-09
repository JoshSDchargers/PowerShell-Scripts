 'About: Clear IE web browser setting to allow SP mapper PS1 script to run.
 'Date Created: 12/3/2019
 'Date Last Modify: 12/3/2019

 
 
 
	Sub IEreset ()
				'IE Reset Automation
		Set objAP = CreateObject("wscript.shell")
			objAP.Run "rundll32.exe inetcpl.cpl ResetIEtoDefaults"
			wscript.sleep 1000

			objAP.AppActivate "Reset Internet Explorer Settings"
			objAP.SendKeys "%r", True
				wscript.sleep 2000

				If objAP.AppActivate("Reset Internet Explorer Settings") Then objAP.SendKeys "%c"
				wscript.sleep 2000

				If objAP.AppActivate("Reset Internet Explorer Settings") Then objAP.SendKeys "%c"
				wscript.sleep 2000

				If objAP.AppActivate("Reset Internet Explorer Settings") Then objAP.SendKeys "%c"

	End Sub


'run cmd commands
	Sub CmdCommands ()
	    Set ShellCommand = CreateObject ("WScript.Shell") 
		ShellCommand.run " net use G: /delete "
		ShellCommand.run " net use H: /delete "
		ShellCommand.run " net use J: /delete "
		ShellCommand.run " net use K: /delete "
		ShellCommand.run " net use L: /delete "
		ShellCommand.run " net use M: /delete "
		ShellCommand.run " net use N: /delete "
		ShellCommand.run " net use O: /delete "
		ShellCommand.run " net use P: /delete "
		ShellCommand.run " net use Q: /delete "
		ShellCommand.run " net use S: /delete "
		ShellCommand.run " net use X: /delete"
		ShellCommand.run "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy ByPass -windowstyle hidden \\CHWNAS\Usr_Dir\SPOD\SPOD_v2.34.ps1"
	End Sub
	
    'Main
	call IEreset()
	call CmdCommands()
	


	
	