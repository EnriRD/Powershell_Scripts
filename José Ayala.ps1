Function Menu 
{
    Clear-Host        
    Do
    {
        Clear-Host 
        Write-Host -Object '' 
        Write-Host -Object 'Shortcut creation '
        Write-Host     -Object '----------------------------'
        Write-Host -Object 'Select Operating System ' -ForegroundColor Yellow
        Write-Host     -Object '----------------------------'
        Write-Host -Object '1.  Windows 8.1 '
        Write-Host -Object ''
        Write-Host -Object '2.  Windows 10 '
        Write-Host -Object ''
        Write-Host -Object '3.  Exit'
        Write-Host -Object $errout
        $Menu = Read-Host -Prompt 'Type a number, and press Enter'
        Switch ($Menu) {
            1 {
            CreateVMSnapshot            
            anyKey
            }
            2 {
                 function gShortcut{
                    param (
                     [string]$BrowserPath = "C:\Program Files\Google\Chrome\Application\chrome.exe",
                     [bool]$ShortcutOnDesktop = $true
                    )
                    $WScriptShell = New-Object -ComObject WScript.Shell
                    if ($ShortcutOnDesktop) {
                         $Shortcut = $WScriptShell.CreateShortcut("$env:USERPROFILE\Desktop\$ShortcutName.lnk") 
                         $Shortcut.TargetPath = $BrowserPath
                         $Shortcut.Arguments = $ShortcutUrl
                         if ($ShortcutIconLocation) {
                             $Shortcut.IconLocation = $ShortcutIconLocation
                         }
                         $Shortcut.Save()
                    }
                    if ($ShortCutInStartMenu) {
                        $Shortcut = $WScriptShell.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk") 
                        $Shortcut.TargetPath = $BrowserPath
                        $Shortcut.Arguments = $ShortcutUrl
                        if ($ShortcutIconLocation) {
                            $Shortcut.IconLocation = $ShortcutIconLocation
                        }
                        $Shortcut.Save()
                    }
                 }
               #Ejemplo "Crea acceso directo con Chrome"
               gshortcut -BrowserPath "C:\Program Files\Google\Chrome\Application\chrome.exe" -ShortcutName "El mundo" -ShortcutUrl "https://www.google.com/" -ShortcutInStartMenu $false -ShortcutIconLocation "C:\favicon\icono.ico"
               [System.Windows.MessageBox]::Show('Shortcut Created Sucessfully')
               }
            3 {
                Exit
            }   
            default {
                $errout = 'Option not valid, try again'
            }
        }
    }
    until ($Menu -eq 'q')
}   
# Ir a Menu
Menu