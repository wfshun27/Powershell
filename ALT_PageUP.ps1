# This script is designed to automate the process of moving from one tab to another in a web browser and refreshing the active page.

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted

while(1 -eq 1){

	$wshell=New-Object -ComObject wscript.shell;
	$wshell.AppActivate('iexplorer'); # Activate on Opera browser
    #$wshell.AppActivate('firefox'); # Activate on Opera browser
	Start-Sleep -Seconds 10  # Interval (in seconds) between switch
    #Sleep 10;  
	$wshell.SendKeys('^{PGUP}'); # Ctrl + Page Up keyboard shortcut to switch tab
    
    #wshell.Sendkeys("%{TAB}")
    Start-Sleep -Seconds 15 
    #Sleep 5
	$wshell.SendKeys('{F5}'); # F5 to refresh active page

}