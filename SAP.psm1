using namespace System.Reflection

$SapGuiPath = "C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"

function Start-SAP {
    [OutputType([System.__ComObject])]
    param (
        [Alias("ConnectionString")] [ValidateNotNullOrWhiteSpace()] [string]$conn,
        [Alias("UserId")] [ValidateNotNullOrWhiteSpace()] [string]$id,
        [Alias("UserPassword")] [ValidateNotNullOrWhiteSpace()] [string]$pass,
        [Alias("ProgramPath")] [string]$path = $SapGuiPath,
        [Alias("NewSession")] [switch]$new,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        if (!(Get-Process saplogon -ErrorAction SilentlyContinue)) {
            Start-Process -FilePath $path -ErrorAction Stop
            Start-Sleep -Seconds 3
        }
        $SapROTWr = New-Object -ComObject "SapROTWr.SapROTWrapper"
        $ready = $false
        for ($i = 0; $i -lt 30; $i++) {
            if ($SapROTWr.GetROTEntry("SAPGUI")) {
                $ready = $true
                break
            }
            Start-Sleep -Seconds 1
        }
        if (!$ready) { throw "SAU GUI connection won't start" }
        $application = Invoke-Method $SapROTWr.GetROTEntry("SAPGUI") "GetScriptingEngine"
        if (!$new) {
            for ($i = 0; $i -lt (Get-Property $application "Connections").Length; $i++) {
                $connection = Get-Property $application "Connections" @($i)
                $session = Get-Property $connection "Sessions" @(0)
                if ($conn -and ((Get-Property $connection "ConnectionString") -match $conn) -and
                    $id -and ((Get-Property (Get-Property $session "Info") "User") -eq $id)) {
                    Invoke-Method (Find-Element $session "wnd[0]") "maximize" | Out-Null
                    return $session
                }
            }
        }
        $connection = Invoke-Method $application "OpenConnectionByConnectionString" @($conn, $true, $true)
        $session = Get-Property $connection "Sessions" @(0)
        Invoke-Method (Find-Element $session "wnd[0]") "maximize" | Out-Null
        Set-Text $session "wnd[0]/usr/txtRSYST-BNAME" $id | Out-Null
        Set-Text $session "wnd[0]/usr/pwdRSYST-BCODE" $pass | Out-Null
        Invoke-Method (Find-Element $session "wnd[0]") "sendVKey" 0 | Out-Null
        Invoke-Method (Find-Element $session "wnd[1]/usr/radMULTI_LOGON_OPT1" -OnErrorContinue) "select" -OnErrorContinue | Out-Null
        Invoke-Method (Find-Element $session "wnd[1]" -OnErrorContinue) "sendVKey" 0  -OnErrorContinue | Out-Null
        Invoke-Method (Find-Element $session "wnd[1]" -OnErrorContinue) "sendVKey" 0  -OnErrorContinue | Out-Null
        Invoke-Method (Find-Element $session "wnd[1]" -OnErrorContinue) "sendVKey" 0  -OnErrorContinue | Out-Null
        Invoke-Method (Find-Element $session "wnd[0]" -OnErrorContinue) "sendVKey" 0  -OnErrorContinue | Out-Null
        Invoke-Method (Find-Element $session "wnd[0]" -OnErrorContinue) "sendVKey" 0  -OnErrorContinue | Out-Null
        Invoke-Method (Find-Element $session "wnd[0]" -OnErrorContinue) "sendVKey" 0  -OnErrorContinue | Out-Null
        if ((Get-Property (Get-Property $session "Info") "User") -eq $id) { return $session }
        else { throw "SAP login '$($id)' password wrong or expired" }
    }
    catch { if ($silent) { return $null } else { throw } }
}

function Stop-SAP {
    [OutputType([bool])]
    param (
        [Alias("SAPSession")] [System.__ComObject]$session,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        if ($session) {
            $connection = Get-Property $session "Parent"
            Invoke-Method $connection "CloseConnection" | Out-Null
        }
        else {
            $application = Invoke-Method (New-Object -ComObject "SapROTWr.SapROTWrapper").GetROTEntry("SAPGUI") "GetScriptingEngine"
            $opened = (Get-Property $application "Connections").Length
            for ($i = 0; $i -lt $opened; $i++) {
                $connection = Get-Property $application "Connections" @(0)
                Invoke-Method $connection "CloseConnection" | Out-Null
            }
            for ($i = 0; $i -lt 10; $i++) { Get-Process saplogon | ForEach-Object { $_.CloseMainWindow() } | Out-Null }
            if (Get-Process saplogon -ErrorAction SilentlyContinue) {
                $shell = New-Object -ComObject WScript.Shell
                $shell.AppActivate((Get-Process saplogon).Id) | Out-Null
                Start-Sleep -Seconds 1
                $shell.SendKeys("%{F4}") | Out-Null
            }
            if (Get-Process saplogon -ErrorAction SilentlyContinue) { Get-Process saplogon | Stop-Process -Force }
            Wait-Process saplogon -Timeout 5 -ErrorAction Stop
        }
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Invoke-SAP {
    param (
        [Alias("SAPObject")] [System.__ComObject]$object,
        [Alias("BindingType")] [System.Reflection.BindingFlags]$type,
        [Alias("InvokeName")] [string]$name,
        [Alias("InvokeArguments")] $arguments
    )
    try {
        $objectType = [System.Type]::GetType($object)
        return $objectType.InvokeMember($name, $type, $null, $object, $arguments)
    }
    catch { throw }
}

function Invoke-Method {
    param (
        [Alias("SAPObject")] [System.__ComObject]$object,
        [Alias("MethodName")] [string]$method,
        [Alias("MethodParameter")] $param,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try { return Invoke-SAP $object InvokeMethod $method $param }
    catch { if ($silent) { return $null } else { throw } }
}

function Get-Property {
    param (
        [Alias("SAPObject")] [System.__ComObject]$object,
        [Alias("PropertyName")] [string]$property,
        [Alias("PropertyParameter")] $param,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try { return Invoke-SAP $object GetProperty $property $param }
    catch { if ($silent) { return $null } else { throw } }
}

function Set-Property {
    param (
        [Alias("SAPObject")] [System.__ComObject]$object,
        [Alias("PropertyName")] [string]$property,
        [Alias("PropertyValue")] $value,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try { return Invoke-SAP $object SetProperty $property $value }
    catch { if ($silent) { return $null } else { throw } }
}

function Find-Element {
    [OutputType([System.__ComObject])]
    param (
        [Alias("SAPObject")] [System.__ComObject]$object,
        [Alias("Element")] [string]$value,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try { return Invoke-Method $object "findById" @($value) }
    catch { if ($silent) { return $null } else { throw } }
}

function Invoke-Click {
    [OutputType([bool])]
    param (
        [Alias("SAPObject")] [System.__ComObject]$object,
        [Alias("Element")] [string]$value,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        Invoke-Method (Find-Element $object $value) "press" | Out-Null
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Set-Text {
    [OutputType([bool])]
    param (
        [Alias("SAPObject")] [System.__ComObject]$object,
        [Alias("Element")] [string]$value,
        [Alias("TextInput")] [string]$text,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        Set-Property (Find-Element $object $value) "text" @($text) | Out-Null
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
}
