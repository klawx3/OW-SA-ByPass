$v = "1.3.0"
$ips_file = "ips.txt"
$bat_name_file = "crear_regla_firewall.bat"
$firewall_rule_name = "Overwatch Bloqueo Sur-America (by Kenshi) v$v"

function Load-Ips($f){    
    try{
        $ips = Get-Content $f -ErrorAction Stop
        return $ips
    } catch {        
        $e = $_.Exception.Message
        $error = "No se ha encontrado el archivo $f (error: $errorMessage)"
        Show-Message($error)
        exit     
    }
}

function Show-Message($message){
    Write-Output $message
    $Shell = New-Object -ComObject "WScript.Shell"
    $Button = $Shell.Popup($message, 0, "", 0)
}

function Clean-IpRange($s){
    $aux = [System.Collections.ArrayList]@()
    foreach($_ip in $s){
        if(!$_ip.StartsWith("#") -and !([string]::IsNullOrEmpty($_ip))){
            $aux += $_ip              
        }
    }
    return $aux
}

function Select-OverwatchFile { #GUI
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.OpenFileDialog
    $browse.filter = "All files (Overwatch.exe)| Overwatch.exe"
    $browse.Title = "Ingrese la ubicación de Overwatch.exe:"
    $browse.ShowDialog() | Out-Null
    $n = $browse.FileName
    if($n){
        return $n
    }else{
        Write-Output "Acciones anuladas..."
        exit
    }
}

function Create-ShorcutWithArg($shorcut_name, $ss){
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\$shorcut_name.lnk")
    $current_path = $(Get-Location).Path
    $Shortcut.Arguments = [string]$ss
    $Shortcut.TargetPath = $current_path + "\" + $bat_name_file
    if ($ss -like "true" -or $ss -eq "1"){
    $Shortcut.IconLocation = $current_path + "\" + "icon\_na.ico"
    }
    if ($ss -like "true" -or $ss -eq "0"){
    $Shortcut.IconLocation = $current_path + "\" + "icon\_las.ico"
    }
    $Shortcut.Save()
}




#main


if($args.Length -eq 0){ # without args - create firewall rule
    $all_ips = Clean-IpRange(Load-Ips($ips_file)) # or die
    $ow_file = Select-OverwatchFile # or die
    $rule = Get-NetFirewallRule -DisplayName $firewall_rule_name -ErrorAction Ignore
    Write-Output "Ruta de overwatch usada: $ow_file"
    try{
        if($rule) {
            Write-Output "La regla de firewall '$firewall_rule_name' existente... intentando sobre-escribir nuevos parametros de ips"       
            Remove-NetFirewallRule -DisplayName $firewall_rule_name -ErrorAction Stop
        }
        Write-Output "Creando regla '$firewall_rule_name'"  
        New-NetFirewallRule -DisplayName $firewall_rule_name -Direction Outbound -Program $ow_file -Action Block -RemoteAddress $all_ips -ErrorAction Stop
    } catch {
        $errorMessage = $_.Exception.Message
        Show-Message("Es probable que no tenga permisisos de administrador (error:$errorMessage)")
        exit
    }
    Write-Output "Creando accesos directos para habilitar y desabilitar la regla de firewall"
    Create-ShorcutWithArg "OW - Habilitar NA" "1"
    Create-ShorcutWithArg "OW - Habilitar LAS" "0"
    Write-Output "Cambios creados con exito!"
} else { #with args
    $enable = $args[0]
    if ($enable -like "true" -or $enable -eq "1") { #enable firewall rule
        Enable-NetFirewallRule -DisplayName $firewall_rule_name -ErrorAction Stop
        Write-Output "Servidores SA bloqueados. OW conectará a NA."
    } elseif ($enable -like "false" -or $enable -eq "0") { #disable firewall rule
        Disable-NetFirewallRule -DisplayName $firewall_rule_name -ErrorAction Stop
        Write-Output "Servidores SA desbloqueados. OW conectará a SA."
    } else { #error
        
    }
}