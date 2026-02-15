param(
    [string]$PrinterIp = '',
    [string]$PrinterName = '',
    [int]$PortNumber = 9100,
    [string]$DriverName = '',
    [switch]$PrintTestPage
)

$ErrorActionPreference = 'Stop'

# Centralized defaults for easy adaptation in other environments.
$Config = @{
    LogDir               = '\\your-fileserver\SharedTools\Add-ClinicPrinter'    # update this
    SmtpHost             = 'smtp.yourdomain.example'                            # update this
    SmtpFallbackIp       = '10.0.0.25'                                          # update this (IP of your SMTPhost incase you can't resolve the hostname)
    EmailFrom            = 'Add-ClinicPrinter@yourdomain.example'               # update this
    EmailTo              = @('it-support@yourdomain.example')                   # update this
    # Optional fallback order used only when -DriverName is not provided.
    AutoDetectDriverCandidates = @(
        'HP Universal Printing PCL 6 (v6.6.0)',
        'HP Universal Printing PCL 6',
        'Microsoft IPP Class Driver',
        'Generic / Text Only'
    )
    PrinterPortTimeoutMs = 2000
    SmtpConnectTimeoutMs = 3000
    SmtpPort             = 25
}

$logDir = $Config.LogDir
if (-not (Test-Path -Path $logDir)) {
    try {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    } catch {
        Write-Host "Unable to access log folder: $logDir"
        Write-Host 'Please confirm you are on your corporate network or VPN.'
        exit 5
    }
}
$logPath = Join-Path $logDir 'Add-ClinicPrinter.log'
$scriptVersion = '1.4.0'
$exitCode = 0
$finalStatus = 'Unknown'
$userEmail = $null
$displayName = $null
$runLines = New-Object System.Collections.Generic.List[string]

function Write-Log {
    param([string]$Message)
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "$timestamp $Message"
    Write-Host $line
    Add-Content -Path $logPath -Value $line
    $script:runLines.Add($line) | Out-Null
}

function Test-PingHost {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName
    )

    # Quick ICMP check to confirm the printer is reachable.
    $result = Test-Connection -ComputerName $ComputerName -Count 2 -Quiet -ErrorAction SilentlyContinue
    return [bool]$result
}

function Get-HostNameFromIp {
    param(
        [Parameter(Mandatory = $true)][string]$IpAddress
    )

    try {
        $dns = Resolve-DnsName -Name $IpAddress -ErrorAction Stop
        $ptr = $dns | Where-Object { $_.Type -eq 'PTR' } | Select-Object -First 1
        if ($ptr -and $ptr.NameHost) {
            return $ptr.NameHost
        }
    } catch {
        return $null
    }

    return $null
}

function Test-TcpPort {
    param(
        [string]$ComputerName,
        [int]$Port,
        [int]$TimeoutMs = $($Config.PrinterPortTimeoutMs)
    )
    # Use a short TCP connect to confirm the printer is reachable on the port.
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $success = $async.AsyncWaitHandle.WaitOne([TimeSpan]::FromMilliseconds($TimeoutMs))
        if (-not $success) {
            return $false
        }
        $client.EndConnect($async) | Out-Null
        return $true
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

function Get-AdUserEmail {
    param(
        [Parameter(Mandatory = $true)][string]$SamAccountName
    )

    # Uses ADSI to avoid requiring the AD module.
    $adsi = [ADSI]"LDAP://RootDSE"
    if (-not $adsi -or -not $adsi.defaultNamingContext) {
        return $null
    }

    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.SearchRoot = "LDAP://$($adsi.defaultNamingContext)"
    $searcher.Filter = "(&(objectClass=user)(sAMAccountName=$SamAccountName))"
    $searcher.PropertiesToLoad.Add("mail") | Out-Null
    $searcher.PropertiesToLoad.Add("userPrincipalName") | Out-Null

    $result = $searcher.FindOne()
    if (-not $result) {
        return $null
    }

    $mail = $null
    if ($result.Properties["mail"] -and $result.Properties["mail"].Count -gt 0) {
        $mail = $result.Properties["mail"][0]
    }

    if (-not $mail -and $result.Properties["userPrincipalName"] -and $result.Properties["userPrincipalName"].Count -gt 0) {
        $upn = $result.Properties["userPrincipalName"][0]
        if ($upn -and $upn -match "@") {
            $mail = $upn
        }
    }

    return $mail
}

function Test-SmtpReachable {
    param(
        [Parameter(Mandatory = $true)][string]$SmtpHost,
        [int]$Port = $($Config.SmtpPort)
    )

    $SmtpHost = $SmtpHost.Trim().Split()[0]

    $result = [ordered]@{
        Host       = $SmtpHost
        Port       = $Port
        CanResolve = $false
        CanConnect = $false
        Reason     = $null
    }

    # If host is an IP address, skip DNS resolution.
    $ipObj = $null
    if ([System.Net.IPAddress]::TryParse($SmtpHost, [ref]$ipObj) -or $SmtpHost -match "\d{1,3}(\.\d{1,3}){3}") {
        $result.CanResolve = $true
    } else {
        try {
            $dns = Resolve-DnsName -Name $SmtpHost -ErrorAction Stop
            if ($dns) {
                $result.CanResolve = $true
            } else {
                $result.Reason = "DNS resolution failed"
                return [PSCustomObject]$result
            }
        } catch {
            $result.Reason = "DNS resolution failed"
            return [PSCustomObject]$result
        }
    }

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($SmtpHost, $Port, $null, $null)
        $wait = $iar.AsyncWaitHandle.WaitOne($Config.SmtpConnectTimeoutMs, $false)
        if ($wait -and $client.Connected) {
            $client.EndConnect($iar)
            $result.CanConnect = $true
        } else {
            $result.Reason = "SMTP port $Port unreachable"
        }
        $client.Close()
    } catch {
        $result.Reason = "SMTP connectivity check failed"
    }

    if (-not $result.CanResolve -and -not $result.Reason) {
        $result.Reason = "DNS resolution failed"
    }
    if ($result.CanResolve -and -not $result.CanConnect -and -not $result.Reason) {
        $result.Reason = "SMTP port $Port unreachable"
    }

    return [PSCustomObject]$result
}

function Send-LogEmail {
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$UserEmail,
        [string]$DisplayName,
        [string[]]$RunLines
    )

    # Best-effort email using the configured SMTP relay.
    $emailServer = $Config.SmtpHost
    $emailServerIp = $Config.SmtpFallbackIp
    $emailFrom = $Config.EmailFrom
    $emailTo = $Config.EmailTo
    $emailCc = @()
    if ($UserEmail) {
        $emailCc += $UserEmail
    }

    $smtpStatus = Test-SmtpReachable -SmtpHost $emailServer -Port $Config.SmtpPort
    if (-not $smtpStatus.CanResolve -or -not $smtpStatus.CanConnect) {
        $smtpStatus = Test-SmtpReachable -SmtpHost $emailServerIp -Port $Config.SmtpPort
        $emailServer = $emailServerIp
    }

    if (-not $smtpStatus.CanResolve -or -not $smtpStatus.CanConnect) {
        $failReason = $smtpStatus.Reason
        if (-not $failReason) { $failReason = "SMTP unreachable" }
        Write-Log "Email not sent: $failReason"
        return
    }

    if (-not $DisplayName) {
        $DisplayName = $env:USERNAME
    }
    $emailSubject = "Add Clinic Printer - $Status - $DisplayName"
    $summary = @(
        "Printer Name: $PrinterName"
        "Printer IP: $PrinterIp"
        "Status: $Status"
        ""
    )
    $emailBody = (($summary + $RunLines) -join "`r`n")

    $prevErrorAction = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'
    try {
        $mailParams = @{
            Body        = $emailBody
            From        = $emailFrom
            To          = $emailTo
            Subject     = $emailSubject
            SmtpServer  = $emailServer
            ErrorAction = 'Stop'
        }
        if ($emailCc -and $emailCc.Count -gt 0) {
            $mailParams.Cc = $emailCc
        }
        Send-MailMessage @mailParams
        Write-Log "Email sent to: $($emailTo -join ', ')"
        if ($emailCc -and $emailCc.Count -gt 0) {
            Write-Log "Email CC: $($emailCc -join ', ')"
        }
    } catch {
        Write-Log ("Email send failed: " + $_.Exception.Message)
    } finally {
        $ErrorActionPreference = $prevErrorAction
    }
}

if ([string]::IsNullOrWhiteSpace($PrinterIp)) {
    $enteredIp = Read-Host 'Enter printer IP address'
    if (-not [string]::IsNullOrWhiteSpace($enteredIp)) {
        $PrinterIp = $enteredIp.Trim()
    }
}

if ([string]::IsNullOrWhiteSpace($PrinterIp)) {
    Write-Log 'No printer IP address provided. Cannot continue.'
    exit 2
}

if ([string]::IsNullOrWhiteSpace($PrinterName)) {
    $enteredName = Read-Host 'Optional printer name (press Enter to skip)'
    if (-not [string]::IsNullOrWhiteSpace($enteredName)) {
        $PrinterName = $enteredName.Trim()
    } else {
        # Use a simple, consistent name if the user skips it.
        $PrinterName = "Printer $PrinterIp"
    }
}

Write-Log "Starting printer install for $PrinterIp"
Write-Log "Script version: $scriptVersion"
Write-Log "User: $env:USERNAME"
Write-Log "Domain: $env:USERDOMAIN"
Write-Log "Computer: $env:COMPUTERNAME"

try {
    $finalStatus = 'Running'

    # Do not request elevation; many end users won't have admin rights.
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

    Write-Log "Pinging $PrinterIp..."
    if (-not (Test-PingHost -ComputerName $PrinterIp)) {
        Write-Log "Cannot reach $PrinterIp. Check the IP address, confirm the printer is on the network, and confirm you're on your corporate network or VPN."
        throw "Ping failed for $PrinterIp."
    }

    $resolvedName = Get-HostNameFromIp -IpAddress $PrinterIp
    if ($resolvedName) {
        Write-Log "DNS PTR: $resolvedName"
    } else {
        Write-Log "DNS PTR: Not found for $PrinterIp"
    }

    # Best-effort Active Directory lookup for user details.
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adUser = Get-ADUser -Identity $env:USERNAME -Properties DisplayName,l,st,mail
        if ($adUser) {
            if ($adUser.DisplayName) {
                $displayName = $adUser.DisplayName
                Write-Log "AD Display Name: $displayName"
            }
            if ($adUser.l -or $adUser.st) { Write-Log "AD City/State: $($adUser.l), $($adUser.st)" }
            if ($adUser.mail) { $userEmail = $adUser.mail }
        }
    } catch {
        Write-Log 'Active Directory lookup not available on this machine.'
    }

    if (-not $userEmail) {
        try {
            $userEmail = Get-AdUserEmail -SamAccountName $env:USERNAME
        } catch {
            $userEmail = $null
        }
    }
    if ($userEmail) {
        Write-Log "User Email: $userEmail"
    } else {
        Write-Log "User Email: Not found"
    }
    if (-not $displayName) {
        $displayName = $env:USERNAME
    }

    if (-not (Test-TcpPort -ComputerName $PrinterIp -Port $PortNumber)) {
        Write-Log "Port $PortNumber did not respond on $PrinterIp. Check network/VPN and try again."
        throw "Printer port $PortNumber did not respond."
    }

    $portName = "IP_$PrinterIp"
    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    if (-not $existingPort) {
        Write-Log "Creating TCP/IP port $portName -> ${PrinterIp}:$PortNumber"
        Add-PrinterPort -Name $portName -PrinterHostAddress $PrinterIp -PortNumber $PortNumber
    } else {
        Write-Log "Port $portName already exists"
    }

    if (-not $DriverName) {
        # Choose a reasonable default driver that is already installed.
        $installedDrivers = Get-PrinterDriver | Select-Object -ExpandProperty Name
        $DriverName = $Config.AutoDetectDriverCandidates | Where-Object { $installedDrivers -contains $_ } | Select-Object -First 1
    }

    if (-not $DriverName) {
        Write-Log 'No suitable printer driver found. Install the vendor driver or specify -DriverName.'
        throw 'No suitable printer driver found.'
    }

    Write-Log "Using driver: $DriverName"

    $existingPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
    if ($existingPrinter) {
        Write-Log "Printer $PrinterName already exists"
        if ($existingPrinter.PortName -ne $portName) {
            Write-Log "Updating port to $portName"
            Set-Printer -Name $PrinterName -PortName $portName
        }
        if ($existingPrinter.DriverName -ne $DriverName) {
            Write-Log "Updating driver to $DriverName"
            Set-Printer -Name $PrinterName -DriverName $DriverName
        }
    } else {
        Write-Log "Adding printer $PrinterName (this can take about 60 seconds)..."
        Add-Printer -Name $PrinterName -PortName $portName -DriverName $DriverName
    }

# Prompt interactively if no switch is provided.
    Write-Log 'Printer added successfully.'
    $finalStatus = 'Success'
} catch {
    Write-Log 'Printer was not added successfully.'
    Write-Log ('Error: ' + $_.Exception.Message)
    $finalStatus = 'Failed'
    $exitCode = 1
} finally {
    Send-LogEmail -Status $finalStatus -UserEmail $userEmail -DisplayName $displayName -RunLines $runLines.ToArray()
}

if ($exitCode -ne 0) {
    Write-Log 'Final status: FAILED.'
}

if ($finalStatus -eq 'Success') {
    if (-not $PrintTestPage) {
        if ($Host.Name -eq 'ConsoleHost') {
            $answer = Read-Host 'Print a test page now? (Y/N)'
            if ($answer -match '^(y|yes)$') {
                $PrintTestPage = $true
            }
        }
    }

    if ($PrintTestPage) {
        Write-Log 'Printing test page...'
        Start-Process -FilePath 'rundll32.exe' -ArgumentList "printui.dll,PrintUIEntry /k /n `"$PrinterName`"" -Wait
    } else {
        Write-Log 'Test page skipped.'
    }
}

if ($exitCode -ne 0) {
    exit $exitCode
}
