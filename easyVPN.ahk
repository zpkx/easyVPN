IfNotExist, config.ini
{
    MsgBox,,Error - VpnHotkey,Missing config.ini file!
    ExitApp
}

IniRead, softTokenShortcut, config.ini, Config, softTokenShortcut
IniRead, vpnClientDir, config.ini, Config, vpnClientDir
IniRead, vpnProfile, config.ini, Config, vpnProfile
IniRead, vpnUsername, config.ini, Config, vpnUsername
IniRead, tokenPin, config.ini, Config, tokenPin
IniRead, connectHotkey, config.ini, Config, connectHotkey
IniRead, disconnectHotkey, config.ini, Config, disconnectHotkey
IniRead, connectOnStartup, config.ini, Config, connectOnStartup, false
IniRead, exitAppAfterConnect, config.ini, Config, exitAppAfterConnect, false
IniRead, gui, config.ini, Config, gui, false

errorExist := false
errorMsg := ""

if (softTokenShortcut = "ERROR" || softTokenShortcut = "")
{
    errorExist := true
    errorMsg := errorMsg . "`nPlease specify a value for softTokenShortcut in config.ini."
}
else IfNotExist, %softTokenShortcut%
{
    errorExist := true
    errorMsg := errorMsg . "`nNo SofToken-II shortcut was found!"
}

if (vpnClientDir = "ERROR" || vpnClientDir = "")
{
    errorExist := true
    errorMsg := errorMsg . "`nPlease specify a value for vpnClientDir in config.ini."
}
else IfNotExist, %vpnClientDir%
{
    errorExist := true
    errorMsg := errorMsg . "`nNo Cisco VPN Client was found!"
}

if (vpnProfile = "ERROR" || vpnProfile = "")
{
    errorExist := true
    errorMsg := errorMsg . "`nPlease specify a value for vpnProfile in config.ini."
}
else IfNotExist, %vpnClientDir%\Profiles\%vpnProfile%.pcf
{
    errorExist := true
    errorMsg := errorMsg . "`nThe VPN profile - " . vpnProfile . " was not found!"
}

if (vpnUsername = "ERROR" || vpnUsername = "")
{
    errorExist := true
    errorMsg := errorMsg . "`nPlease specify a value for vpnUsername in config.ini."
}

if (tokenPin = "ERROR" || tokenPin = "")
{
    errorExist := true
    errorMsg := errorMsg . "`nPlease specify a value for tokenPin in config.ini."
}

if (connectHotkey = "ERROR" || connectHotkey = "")
{
    errorExist := true
    errorMsg := errorMsg . "`nPlease specify a value for connectHotkey in config.ini."
}

if (disconnectHotkey = "ERROR" || disconnectHotkey = "")
{
    errorExist := true
    errorMsg := errorMsg . "`nPlease specify a value for disconnectHotkey in config.ini."
}

if (errorExist = true)
{
    MsgBox,,Error - VpnHotkey,Config Errors:`n%errorMsg%
    ExitApp
}

hotkey, %connectHotkey%, connect
hotkey, %disconnectHotkey%, disconnect

if (connectOnStartup = "false")
{
    return
}


; Connection to VPN
connect:

; If already connected to VPN, then popup a warning.
statFile = %A_Temp%\vpnStat.tmp
runWait, %comspec% /c "%vpnClientDir%\vpnclient.exe" stat tunnel > %statFile%,,Hide
FileReadLine, line, %statFile%, 8
;MsgBox %line%
FileDelete, %statFile%
IfInString, line, Connection Entry
{
    MsgBox,,Warning - VpnHotkey, Your VPN is already connected.`n`n%line%
    return
}

run, "%softTokenShortcut%" ; Run the shortcut of SofToken
WinWait, SofToken II
WinActivate, SofToken II
sendinput, %tokenPin%{enter} ; Input PIN code
sleep, 200
sendinput, {tab 2}{Ctrl Down}{c}{Ctrl Up}
password = %Clipboard%
WinClose

if (gui = "true")
{
    ; GUI
    run, "%vpnClientDir%\vpngui.exe" -c -sd -user %vpnUsername% -pwd %password% "%vpnProfile%"
    WinWait, VPN Client,,3 ;If error occurs, then return
    if (ErrorLevel = 0)
    {
        return
    }
    winTile := "VPN Client  |  Banner"
    WinWait,%winTile%,,50
    if ErrorLevel
    {
        MsgBox,,Warning - VpnHotkey, Script timeout!, 5
        return
    }
    WinActivate,%winTile%
    sendinput, {enter}
}
else
{
    ; CLI
    run, %comspec% /c ""%vpnClientDir%\vpnclient.exe" connect "%vpnProfile%" user %vpnUsername% pwd %password%"
}

if (exitAppAfterConnect = "true")
{
    ExitApp
}
return


; Disconnect from VPN
disconnect:
run, %comspec% /c "%vpnClientDir%\vpnclient.exe" disconnect
return
