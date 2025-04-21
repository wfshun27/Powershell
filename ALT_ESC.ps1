#If the registry entry is missing, you can use #ALT_ESC as an alternative.

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted

while(1 -eq 1){

	$wshell=New-Object -ComObject wscript.shell; # Create a new WScript.Shell object
	$wshell.AppActivate('iexplorer'); # Activate on Opera browser
    #$wshell.AppActivate('firefox'); # Activate on Opera browser
	Start-Sleep -Seconds 10  # Interval (in seconds) between switch
	wshell.Sendkeys("%{ESC}") # Alt + Esc keyboard shortcut to switch tab
    Start-Sleep -Seconds 10 
	$wshell.SendKeys('{F5}'); # F5 to refresh active page
}