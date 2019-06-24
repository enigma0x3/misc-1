#New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR
$DcomObjects = gwmi -Class win32_dcomapplicationsetting
foreach($DcomObj in $DcomObjects)
{
    $Guid = $Type = $Instance = $null

    try {
        $Guid = $DcomObj.AppId.Replace('{','').Replace('}','')
        $Type = [Type]::GetTypeFromCLSID([guid]$Guid, ‘localhost’)
        $Instance = [Activator]::CreateInstance($Type)
        
        "$($DcomObj.Caption) - $($DcomObj.Description) - $($DcomObj.AppId)" | Out-File -Append -Encoding ascii .\com_out.txt
        $Members = $Instance | gm
        $Members | ?{$_.Name -notmatch 'CreateObjRef|Equals|GetHashCode|GetLifetimeService|GetType|GetType|ToString|InitializeLifetimeService'} | ft -a | Out-File -Append -Encoding ascii .\com_out.txt
        

    } catch {
        "$($DcomObj.Caption) - $($DcomObj.Description) - $($DcomObj.AppId)" | Out-File -Append -Encoding ascii .\com_errors.txt
        $_ | Out-String -Width 400 | Out-File -Append -Encoding ascii .\com_errors.txt
    }

    if($Instance)
    {
        $null = [Runtime.Interopservices.Marshal]::ReleaseComObject($Instance)
    }
}
