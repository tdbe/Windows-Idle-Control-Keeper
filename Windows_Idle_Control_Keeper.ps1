<#
.SYNOPSIS
	MIT License, Copyright (c) 2026 Tudor Berechet [tdbe](https://github.com/tdbe) 

	# Intro:
	
	[Saul Goodman voice] Can't get reliable sleep? Feeling like it's out of your control? Well fret not! Just run this script and start counting your sheep! 
		WICK - Windows Idle Control Keeper
	
	This script detects Idle activity with your specific thresholds and conditions, and triggers and/or prevents Windows Sleep on Your terms. Detects activity on CPU, network, storage, mouse, and peak sound value, instances per time period, to determine if an Idle timer should continue or be broken. It does not affect and is not affected by (auto) windows screen locking, or (auto) turning off the display.
	
	I don't usually post my system scripts but it annoyed me that for such a wide need, there was nothing out there but forum threads of people using ancient and partial tools like (DontSleep!.exe)[https://www.softwareok.com/?Download=DontSleep] (from 2014)[https://www.chip.de/downloads/Don-t-Sleep_42626965.html]
	
	# Features, Dependencies, Log Example, Notes, Parameters:

	## Features:

	- Does not require administrator permission.
	- Works even if windows is locked. Also works if logged out or never logged in (if you start it at system start via task scheduler).
	- Shows warning / abort window for $SleepAbortWindowCountdownSeconds before triggering a sleep (if in an interactive session (not locked or logged out)).
	- Dynamically reads (every minute) from your currently active windows power plan (plugged in or battery) to check sleep and also hibernate times.
	- Determines idle by accumulating sustained activity events over certain timeframes, through checks every second: based on if there's CPU, Network, sorage (without waking sleeping hard disks), audio, and mouse activity.
	- Can prevent windows from sleeping until this script decides it's time to sleep.
	- Can set a sleep time for longer than 5h (the max that Windows power plan allows for some gormless reason).
	- Allows a blacklist for logical drives e.g. `"L", "A", "N"` - you may have drives that have activity you consider passive and you're okay sleeping on. But also keep in mind the NetworkThresholdKBps setting.
	- Can be paused while running by creating a `.ignore_running_Windows_Idle_Control_Keeper_script` flag file (e.g. renamed empty txt file).
	- Logs what's going on, at LogPath, so you know at what time Idle state was broken and after how much idle time. (or if there were errors) (log cleans itself up to stay less than LogMaxSizeMB)
	- It does not affect and is not affected by (auto) windows screen locking, or (auto) turning off the display.

	## Dependencies:
		- python (and the `checkIfAudioIsPlaying.py` script, which requires `pip install pycaw`)
		- virtually any .net (C# capability) installed on the system (for `SetThreadExecutionState`)
		- powershell 5.1 (the latest is powershell 7+)
		- You need to check that the paths are correct / to your liking. Set the corresponding "*Path" parameters.

	## Log Example:
	
	```
	[2026-04-30 00:11:58] [INFO] "~*------- W.I.C.K. started. -------"
	[2026-04-30 00:11:58] [INFO]   Log path: C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log
	[2026-04-30 00:11:58] [INFO] Power plan idle timeout: 60 min: sleep=3600 sec, hibernate=3600 sec
	[2026-04-30 00:11:58] [INFO]   Using windows power plan's minimum(sleep, hibernate) value as the idle timeout
	(checks the current active power plan value every: 60 sec)
	[2026-04-30 00:11:58] [INFO] It's been 171368.983134705 minute(s) since the last update, which means we
	were sleeping or somehow lagging a lot, Resetting idle counter.
	[2026-04-30 00:13:38] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:15:49] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:18:03] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:20:18] [INFO] [IDLE BREAKER] Sustained audio playing for 5 sec, Resetting idle counter.
	[2026-04-30 00:22:27] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:24:35] [INFO] [IDLE BREAKER] Network: 3/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:26:53] [INFO] [IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:29:07] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:31:16] [INFO] [IDLE BREAKER] Sustained audio playing for 5 sec, Resetting idle counter.
	[2026-04-30 00:33:22] [INFO] [IDLE BREAKER] Disk: 4/6 samples > 1250 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:35:41] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:38:00] [INFO] [IDLE BREAKER] Network: 3/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:42:30] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:44:38] [INFO] [IDLE BREAKER] Mouse moved 160.6 px > 10, Resetting idle counter.
	[2026-04-30 00:46:52] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:49:18] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:51:39] [INFO] [IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:53:57] [INFO] [IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:53:59] [INFO] CPU: 1 % | Disk: 8279 KBps | Net: 1201 KBps | MouseDelta: 0 px | Idle: 0 min / 60 min
	```

	## Notes: 
	
	### Note: I've always nuked Modern Standby from every PC I touched, because we have literally 0.0f low-power hardware and protocol standards, and I don't want constant 100W power draw, and for laptops my battery to run out in 2 hours while "sleeping" with the lid closed (Microsoft is the most infuriating thing in the history of ever). You can have a look at how I printed and fetched the sleep AC/DC settings, and figure out the parsing of S3 yourself if you want. PRs welcome.
	
	### Note: I don't vibecode anything I consider even remotely reliable, and either way this script is read-through and tested. But for this project I did try out LLMs, otherwise I wouldn't be caught dead writing 700 lines of powershell script of all things. I used qwen 3 coder next 80b a3b q6, qwen 3.6 35b a3b q8, and qwen 3.6 27b q4; they're "great" (within 5-10% of the huge frontier models) but simultaneously also completely shit at even such a simple job, and not just because this solution doesn't already exist: ie they picked network and storage checks that take at least 1s to return a value, and were calling them repeatedly in loops per disk and per adapter, resulting in a while loop that runs once every 7-10s.. So the verdict is I had to do all the thinking myself. It only oneshotted the logging, the cpu, the sleep functions, an dthe .PARAM list. Also the audio checking I had to research and write myself in python after many wildly off LLM solutions.
	
	## Parameters:

	Parameters you can set when calling the script or adding it to Task Scheduler (the ones you skip will have defaults).
	To see all parameters and their description, run this command: `get-help ./windows_idle_control_keeper.ps1 -detailed`
	
	e.g. ./windows_idle_control_keeper.ps1 -FollowTheSameSleepTimeSettingAsYourPowerPlan $true -OnlyThisScriptCanCauseWindowsToSleep $true -FollowTheSameSleepTimeSettingAsYourPowerPlan $true -FallbackIdleMinutes 30 -etc. -etc.
	
	or
	
	e.g. ./windows_idle_control_keeper.ps1 -FollowTheSameSleepTimeSettingAsYourPowerPlan $false -OnlyThisScriptCanCauseWindowsToSleep $true -FollowTheSameSleepTimeSettingAsYourPowerPlan $true -FallbackIdleMinutes 30 -IdleDurationMinutes 720 -etc. -etc.

.PARAMETER OnlyThisScriptCanCauseWindowsToSleep
  Uses `SetThreadExecutionState` to prevent other processes from overriding the sleep state.  
  Ensures only this script can trigger sleep. (default: $true)
  
.PARAMETER FollowTheSameSleepTimeSettingAsYourPowerPlan
  Read idle timeout from Windows power plan (default: $true)

.PARAMETER FallbackIdleMinutes
  Fallback idle minutes if power plan is disabled (default: 30)

.PARAMETER IdleDurationMinutes
  Use this to set more than the weird 5h max limit that Windows power plan lets you set.  
  Ignored if FollowTheSameSleepTimeSettingAsYourPowerPlan == $true (default: 720) (12 hours)

.PARAMETER CpuThresholdPercent
  CPU usage above this resets idle timer (default: 6)

.PARAMETER DiskThresholdKBps
  Disk I/O (KB/s) above this resets idle timer (default: 1250)

.PARAMETER NetworkThresholdKBps
  Network I/O (KB/s) above this resets idle timer (default: 850)

.PARAMETER MouseThresholdPixels
  Mouse movement above this many pixels resets idle timer (default: 10)

.PARAMETER ActiveSamplesWithinInterval
  How many instances (seconds) of activity must be detected within the last $ActivityDetectionPeriod seconds for us to consider that activity an idle breaker (default: 3)

.PARAMETER ActivityDetectionPeriod
  Seconds window to check for sustained activity (default: 6)

.PARAMETER ActivityDetectionPeriodAudio
  Separate timeout for audio - counts if there was constant sound in this last period of seconds (default: 5)

.PARAMETER DiskBlacklistDrives
  Allows blacklist for logical drives e.g. `"L", "A", "N"` - drives that have activity but you consider passive and you're okay sleeping on them.  
  Keep in mind the NetworkThresholdKBps setting.  
  (default: @("E", "F"))

.PARAMETER PycawAudioCheckerPath
  Full path to the Python script used to detect audio playback.  
  (default: "C:\Commands_And_Logs\Pycaw_check_if_audio_is_playing.py")

.PARAMETER PythonPath
  Full path to the Python executable used to run the audio checker script.
  (default: "$env:USERPROFILE\AppData\Local\Programs\Python\Python312\python.exe")

.PARAMETER PauseFlagPath
  Path to a flag file. If this file exists, the script pauses monitoring and skips sleep.  
  Allows manual pause/resume by creating/deleting the file.
  (default: "C:\Command_And_Logs\.ignore_running_Windows_Idle_Control_Keeper_script")

.PARAMETER LogPath
  Full path to the log file (default: "C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log")

.PARAMETER LogMaxAgeDays
  Keep logs this many days (default: 30)

.PARAMETER LogMaxSizeMB
  Rotate log if larger than this MB (default: 10)

.PARAMETER logToFileIntervalSeconds
  To prevent writing to file every second while you're using the PC, it won't log unless it's been idle for this many seconds. (default: 60)
  
.PARAMETER LogToConsoleVerbose
  Whether to log to the console (not log file) as often as there is an event in the constant loop (default: true)

.PARAMETER SleepAbortWindowCountdownSeconds
  Seconds to show the sleep abort dialog before triggering sleep (default: 60)

.PARAMETER SampleIntervalSec
  How often to sample system metrics (default: 1)
  
.PARAMETER powerPlanPollIntervalSeconds
  Dynamically reads from your currently active windows power plan (plugged in or battery) to check sleep and also hibernate times. (default: 60)
#>

param(
    [bool]$OnlyThisScriptCanCauseWindowsToSleep = $true, # uses SetThreadExecutionState
    [bool]$FollowTheSameSleepTimeSettingAsYourPowerPlan = $true, # polled every minute so it knows if you switched to battery or if the power plan changed.
    [int]$FallbackIdleMinutes = 30,
    [int]$IdleDurationMinutes = 60*12, # Using this you can set more than the weird 5h max limit that windows power plan lets you set. Ignored if FollowTheSameSleepTimeSettingAsYourPowerPlan == $true
    [double]$CpuThresholdPercent = 6,
    [long]$DiskThresholdKBps = 1250,
    [long]$NetworkThresholdKBps = 850,
    [int]$MouseThresholdPixels = 10, # If mouse moved more than this number of pixels.
    [int]$ActivityDetectionPeriod = 6, # we determine if idle was broken, if enough activity existed in the last $ActivityDetectionPeriod seconds.
	[int]$ActiveSamplesWithinInterval = 3, # how many instances (seconds) of activity must be detected within the last $ActivityDetectionPeriod seconds for us to consider that activity an idle breaker.
	[int]$ActivityDetectionPeriodAudio = 5, # Separate timeout for audio, this one just simply counts if there was constant sound in this last period of seconds. This way a windows alert sound doesn't break idle.
	[string[]]$DiskBlacklistDrives = @("E", "F"),
    [string]$PycawAudioCheckerPath = "C:\Commands_And_Logs\Pycaw_check_if_audio_is_playing.py",
    [string]$PythonPath = "$env:USERPROFILE\AppData\Local\Programs\Python\Python312\python.exe",
	[string]$LogPath = "C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log",
    [string]$PauseFlagPath = "C:\Commands_And_Logs\.ignore_running_Windows_Idle_Control_Keeper_script",
    [int]$LogMaxAgeDays = 30,
    [int]$LogMaxSizeMB = 10,
	[int]$logToFileIntervalSeconds = 60, # To prevent writing to file every second while you're using the PC, it won't log unless it's been idle for this many seconds.
	[bool]$LogToConsoleVerbose = $true, # Whether to log to the console (not log file) as often as there is an event in the constant loop
    [int]$SleepAbortWindowCountdownSeconds = 60,
    [int]$SampleIntervalSec = 1,
	[int]$powerPlanPollIntervalSeconds = 60
)

$myUnixTimeEpochStart = Get-Date '2026-01-01'
$minutesPassedLastFrame = 0

# NOTE: / TODO: If you ever want to lock the pc on your own terms before windows does, you can use this:
# Add-Type -MemberDefinition "[DllImport(\"user32.dll\")] public static extern bool LockWorkStation();" -Name LockScreen -Namespace NativeMethods
# [NativeMethods.LockScreen]::LockWorkStation()

$typeName = 'WindowsSleepWrangler'
if ($null -eq ([type]$typeName -as [type])) {
	$code = @"
		using System;
		using System.Runtime.InteropServices;
		public class WindowsSleepWrangler {
			[DllImport("kernel32.dll", SetLastError = true)]
			public static extern uint SetThreadExecutionState(uint esFlags);

			public static void IgnoreIdleTimers() {
				// 0x80000000 ES_CONTINUOUS (keep this command active until we call ES_CONTINUOUS again with some other flags) | 
				// 0x00000001 ES_SYSTEM_REQUIRED (don't sleep) | 
				// 0x00000002 ES_DISPLAY_REQUIRED (don't turn off display)
				SetThreadExecutionState(0x80000000 | 0x00000001);// | 0x00000002);
			}

			public static void StopIgnoringIdleTimers() {
				// 0x80000000 ES_CONTINUOUS (if not accompanied by other flags, it keeps this default state until we call ES_CONTINUOUS again with some other flags
				SetThreadExecutionState(0x80000000);
			}
		}
"@

	Add-Type -TypeDefinition $code -Language CSharp
}

# plugged in means AC power in power plan, battery means DC power. Important for sleep timers (different per AC / DC)
function IsComputerPluggedIn{
	return (Get-WmiObject -Class BatteryStatus -Namespace root\wmi).PowerOnLine
}

# --- Helper: Read idle timeout from power plan (in minutes) ---
function Get-PowerPlanIdleTimeoutMinutes {
    try {
        # Get active power plan GUID
		$activePlan = (powercfg /getactivescheme) -replace '^([0-9a-f-]+).*', '$1'
        if (-not $activePlan) { throw "Failed to get active power plan" }
		
		$activePlan = $activePlan.Substring(19, 36).Trim()
		
		$isAC = IsComputerPluggedIn
		
		#Write-Log "active power plan: $activePlan" "INFO"
		# Power Scheme GUID: 381b4222-f694-41f0-9685-ff5bb260df2e  (Balanced)
		
		# Subgroup GUID: 238c9fa8-0aad-41ed-83f4-97be242c8f20  (Sleep)
		# GUID Alias: SUB_SLEEP
		$SUB_SLEEP = "238c9fa8-0aad-41ed-83f4-97be242c8f20"
		
		# Power Setting GUID: 29f6c1db-86da-48c5-9fdb-f2b67b1f44da  (Sleep after)
		# GUID Alias: STANDBYIDLE
		$STANDBYIDLE = "29f6c1db-86da-48c5-9fdb-f2b67b1f44da"
		
        # Read StandbyIdleTimeout (AC)
        $sleepOutput = powercfg /q $activePlan $SUB_SLEEP $STANDBYIDLETIMOUT 2>$null
        if (-not $sleepOutput) { throw "powercfg failed for standby" }

		#Write-Log "sleepOutput: $sleepOutput" "INFO"

        $sleepSec = 0
		# Read StandbyIdleTimeout (AC/DC)
		$powerSettingIndex = if ($isAC) {
			'Current AC Power Setting Index'
		} else {
			'Current DC Power Setting Index'
		}
		$regex = [regex]("$powerSettingIndex" + ":\s+(0x[0-9a-fA-F]+)")
		$match = $regex.Match($sleepOutput)

        #if ($sleepOutput -match 'Current AC Power Setting Index:\s+(\d+)') {
			#$sleepSec = [int]$matches[1]
        if ($match.Success) {
			$hexValue = $match.Groups[1].Value
			#Write-Log "hexValue: $hexValue" "INFO"
			# Convert hex to decimal: "0x00003840" → 14400
			$sleepSec = [int]::Parse($hexValue.Substring(2), 'AllowHexSpecifier')
		}

		#Write-Log "polled sleepSec: $sleepSec" "INFO"
		Write-Host-Wrapper "polled sleepSec: $sleepSec" "INFO"

		# Power Setting GUID: 9d7815a6-7ee4-497e-8888-515a05f02364  (Hibernate after)
		# GUID Alias: HIBERNATEIDLE
		$HIBERNATEIDLETIMOUT = "9d7815a6-7ee4-497e-8888-515a05f02364"
		
		$hibernateSec = 0
        # Read HibernateIdleTimeout
        $hibernateOutput = powercfg /q $activePlan $SUB_SLEEP $HIBERNATEIDLETIMOUT 2>$null
        if (-not $hibernateOutput) { 
			throw "powercfg failed for hibernate" 
		} else {
			$regex = [regex]'Current AC Power Setting Index:\s+(0x[0-9a-fA-F]+)'
			$match = $regex.Match($sleepOutput)
			
			#if ($hibernateOutput -match 'Current AC Power Setting Index:\s+(\d+)') {
				#$hibernateSec = [int]$matches[1]
			if ($match.Success) {
				$hexValue = $match.Groups[1].Value
				$hibernateSec = [int]::Parse($hexValue.Substring(2), 'AllowHexSpecifier')
			}

			#Write-Log "polled hibernateSec: $hibernateSec" "INFO"
			Write-Host-Wrapper "polled hibernateSec: $hibernateSec" "INFO"
		}
		
        # Determine effective timeout (min of sleep/hibernate, or 0 if disabled)
        $effectiveSec = 0
        if ($sleepSec -gt 0 -and $hibernateSec -gt 0) {
            $effectiveSec = [math]::Min($sleepSec, $hibernateSec)
        } elseif ($sleepSec -gt 0) {
            $effectiveSec = $sleepSec
        } elseif ($hibernateSec -gt 0) {
            $effectiveSec = $hibernateSec
        }
		
        # Convert to minutes, fallback if disabled
        if ($effectiveSec -le 0) {
            Write-Log "Power plan idle timeout: disabled, using fallback: $FallbackIdleMinutes min" "INFO"
			Write-Log "effectiveSec: $FallbackIdleMinutes" "INFO"
            return $FallbackIdleMinutes
        }

		#Write-Log "effectiveSec: $effectiveSec" "INFO"
		Write-Host-Wrapper "effectiveSec: $effectiveSec" "INFO"

        $minutes = [math]::Round($effectiveSec / 60)
        if($IdleDurationMinutes -ne $minutes) {
			Write-Log "Power plan idle timeout: $minutes min: sleep=$sleepSec sec, hibernate=$hibernateSec sec" "INFO"
		}
        return $minutes
    }
    catch {
        Write-Log "Failed to read power plan: $_, using fallback: $FallbackIdleMinutes min" "WARN"
        return $FallbackIdleMinutes
    }
}

# --- Logging Setup (unchanged) ---
$script:LogDir = Split-Path $LogPath -Parent
if (-not (Test-Path $script:LogDir)) {
    try {
        New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null
        Write-Host-Wrapper "Created log directory: $script:LogDir" "INFO"
    }
    catch {
        Write-Host-Wrapper "ERROR: Could not create log directory: $script:LogDir. Logging disabled." "ERROR"
        $LogPath = $null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    if (-not $LogPath) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    Write-Host "[(logged)] $logEntry"

    try {
        Add-Content -Path $LogPath -Value $logEntry -ErrorAction Stop

        $file = Get-Item $LogPath -ErrorAction SilentlyContinue
        if ($file -and $file.Length -gt ($LogMaxSizeMB * 1MB)) {
            Rotate-Log
        }
    }
    catch {
        Write-Host-Wrapper "WARNING: Could not write to log file: $_" "WARN"
    }
}

function Write-Host-Wrapper {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    Write-Host "[(not logged)] $logEntry"
}

function Rotate-Log {
    param([string]$Path = $LogPath)
    if (-not (Test-Path $Path)) { return }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = "$Path.$timestamp"
    try {
        Move-Item -Path $Path -Destination $backupPath -Force
        Write-Log "Rotated log to: $backupPath" "INFO"
    }
    catch {
        Write-Log "Failed to rotate log: $_" "ERROR"
    }

    Get-ChildItem $script:LogDir -Filter "IdleSleepMonitor.log.*" -File |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogMaxAgeDays) } |
        Remove-Item -Force
}

# --- Detect if interactive session ---
# [System.Environment]::UserInteractive is true even when the computer is locked. It's only false if you never logged in, or if it's a specific non interactive system account.
function Test-IsInteractiveSession {
    try {
		$isInteractive = [System.Environment]::UserInteractive
		#Write-Log "interactive???? $isInteractive" "INFO"
        return $isInteractive
    }
    catch {
		Write-Log "Could not query System.Environment::UserInteractive" "ERROR"
        return $false
    }
}

# --- Sleep API ---
function Enter-SleepState {
    Add-Type -TypeDefinition @'
    using System;
    using System.Runtime.InteropServices;
    public static class PowerManagement {
        [DllImport("powrprof.dll", SetLastError = true)]
        public static extern bool SetSuspendState(bool hiberate, bool forceCritical, bool disableWakeEvent);
    }
'@
    $result = [PowerManagement]::SetSuspendState($false, $true, $false)
    if (-not $result) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Log "Sleep failed! Win32 error: $err" "ERROR"
    }
}

# --- CPU ---
function Get-CpuUsagePercent {
    #(Get-WmiObject Win32_Processor | Measure-Object LoadPercentage -Average).Average
	$avg = (Get-WmiObject Win32_Processor | Measure-Object LoadPercentage -Average).Average
    if ($null -eq $avg) { 
		return 0.0 
	}
    return [double]$avg
}

# --- Disk: SAFE - uses LogicalDisk counters (no spin-up risk) ---
function Get-DiskIoKBps {
    try {
        # Call once - get all formatted disk counters
        $diskCounters = Get-CimInstance Win32_PerfFormattedData_PerfDisk_LogicalDisk -ErrorAction Stop
        
        # Filter out "_Total" and blacklisted drives
        $allowedDrives = $diskCounters | ForEach-Object {
            if ($_.Name -eq "_Total") { return }
            $driveLetter = $_.Name -replace ':', ''
            if ($DiskBlacklistDrives -contains $driveLetter) { return }
            $_  # Pass through if allowed
        }

        if (-not $allowedDrives) { return 0 }

        $totalReadKBps = 0
        $totalWriteKBps = 0

        foreach ($disk in $allowedDrives) {
            # Win32_PerfFormattedData_PerfDisk_LogicalDisk properties (bytes/sec)
			$readBytes = $disk.DiskReadBytesPerSec
            $writeBytes = $disk.DiskWriteBytesPerSec
			if ($null -eq $readBytes) { 
				$readBytes = 0 
			}
			if ($null -eq $writeBytes) { 
				$writeBytes = 0 
			}            
            # Convert to KB/s
            $totalReadKBps += $readBytes / 1KB
            $totalWriteKBps += $writeBytes / 1KB
        }

        return [long]($totalReadKBps + $totalWriteKBps)
    }
    catch {
        Write-Log "Get-DiskIoKBps failed: $_" "WARN"
        return 0
    }
}

# --- Network: Sees all active non-virtual network interfaces. ---
function Get-NetworkIoKBps {
    try {
        $netCounters = Get-CimInstance Win32_PerfFormattedData_Tcpip_NetworkInterface -ErrorAction Stop
        
        # Filter out loopback/internal adapters (keep only active interfaces)
        $activeInterfaces = $netCounters | Where-Object {
            $_.Name -ne "_Total" -and 
            $_.Name -notmatch 'Loopback|Teredo|isatap|6to4'  # Exclude tunnel adapters
        }

        if (-not $activeInterfaces) { return 0 }

        $totalRxBytes = 0
        $totalTxBytes = 0

        foreach ($iface in $activeInterfaces) {
            # Win32_PerfFormattedData_Tcpip_NetworkInterface properties (bytes/sec):
            $rxBytes = $iface.BytesReceivedPersec
            $txBytes = $iface.BytesSentPersec
			if ($null -eq $rxBytes) { 
				$rxBytes = 0 
			}
			if ($null -eq $txBytes) { 
				$txBytes = 0 
			} 
            
            $totalRxBytes += $rxBytes
            $totalTxBytes += $txBytes
        }

        # Convert to KB/s (bytes → KB)
        return [long](($totalRxBytes + $totalTxBytes) / 1KB)
    }
    catch {
        Write-Log "Get-NetworkIoKBps failed: $_" "WARN"
        return 0
    }
}

# --- Mouse movement check (interactive only) ---
function Get-MouseMovementPixels {
	$retIsInteractiveSession = Test-IsInteractiveSession
    if (-not $retIsInteractiveSession) {
        Write-Log "Mouse check skipped because non-interactive session." "DEBUG"
        return [PSCustomObject]@{ X = 0; Y = 0 }
    }

    if (-not ("System.Windows.Forms.Cursor" -as [type])) {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Log "Failed to load WinForms: $_" "WARN"
            return [PSCustomObject]@{ X = 0; Y = 0 }
        }
    }

    try {
        $pos = [System.Windows.Forms.Cursor]::Position
        return [PSCustomObject]@{
            X = $pos.X
            Y = $pos.Y
        }
    }
    catch {
        Write-Log "Failed to get mouse position: $_" "WARN"
        return [PSCustomObject]@{ X = 0; Y = 0 }
    }
}

# --- Sound: Detects if any audio is playing ---
function Get-AudioIsPlaying {
    if (-not [System.Environment]::UserInteractive) { return $false }

    # Capture stdout only, force case-insensitive comparison, guarantee boolean
	$rawOutput = & $PythonPath $PycawAudioCheckerPath 2>$null | ForEach-Object { $_.Trim().ToLower() }
	$isPlaying = $rawOutput -eq 'true'
	#Write-Log "isPlaying: $isPlaying" "INFO"
	return $isPlaying
}

# --- Show abort dialog (interactive only) ---
function Show-AbortDialog {
    param([int]$Seconds)
    $retIsInteractiveSession = Test-IsInteractiveSession
    if (-not $retIsInteractiveSession) {
        Write-Log "Non-interactive session - skipping message box - no user session." "DEBUG"
        return $false
    }

    $title = "System Idle Sleep Warning"
    $msg = "Your PC is idle for $IdleDurationMinutes minutes.`n`n" +
           "It will sleep in $Seconds seconds.`n" +
           "Press OK to abort."

    $wsh = New-Object -ComObject WScript.Shell
    $result = $wsh.Popup($msg, $Seconds, $title, 4 + 32)

    return ($result -eq 1)
}

# --- Main Loop ---
Write-Log "~*------- W.I.C.K. started. -------"
Write-Log "  Log path: $LogPath"
#Write-Log "  Dynamic idle timeout (checking the current active power plan value every: $powerPlanPollIntervalSeconds)"
if ($FollowTheSameSleepTimeSettingAsYourPowerPlan) {
    $IdleDurationMinutes = Get-PowerPlanIdleTimeoutMinutes
    Write-Log "  Using windows power plan's minimum(sleep, hibernate) value as the idle timeout: $IdleDurationMinutes min. (Checks the current active power plan value every: $powerPlanPollIntervalSeconds sec.)"
} else {
    Write-Log "  Using manual idle timeout value: system considered idle at: $IdleDurationMinutes min."
}

# Load WinForms once (for mouse check)
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
}
catch {
    Write-Log "Failed to load WinForms: $_" "WARN"
}

# Get initial mouse position
$prevMouse = Get-MouseMovementPixels

# Sliding windows for sustained detection
$maxSamples = [int]([math]::Ceiling($ActivityDetectionPeriod / $SampleIntervalSec))
$maxSamplesAudio = [int]([math]::Ceiling($ActivityDetectionPeriodAudio / $SampleIntervalSec))
$idleSeconds = 0
$cpuHistory = New-Object 'System.Collections.Queue' $maxSamples
$diskHistory = New-Object 'System.Collections.Queue' $maxSamples
$netHistory = New-Object 'System.Collections.Queue' $maxSamples
$audioHistory = New-Object 'System.Collections.Queue' $maxSamplesAudio

$nextPowerPlanPoll = $powerPlanPollIntervalSeconds

if (Test-Path $PauseFlagPath) {
	Write-Log "Script pause flag file is present - while loop running but skipping until flag removed or renamed." "INFO"
}

# NOTE: if this script is frozen or closed unexpectedly in a way that [WindowsSleepWrangler]::StopIgnoringIdleTimers() doesn't get called, windows may not go to sleep again until it is called or until it's restarted. But since we use try - finally, it should auto clean up after itself unless you somehow freeze the thread.
if($OnlyThisScriptCanCauseWindowsToSleep -eq $true){
	[WindowsSleepWrangler]::IgnoreIdleTimers()
}

try {
	while ($true) {
		#Write-Log "Tick. idleSeconds: $idleSeconds" "INFO"
		#Write-Host-Wrapper "Tick. idleSeconds: $idleSeconds" "INFO"
		
		# turns out we don't need to wait because some of these commands have a 1s delay. Would be nice to guarantee a time schedule, but we don't really care, we just poll roughly, and we account for long lag spikes.
		Start-Sleep -Milliseconds (10)
		
		if (Test-Path $PauseFlagPath) {
			$idleSeconds = 0
			continue
		}
		
		$minutesPassed = (Get-Date).Subtract($myUnixTimeEpochStart).TotalMinutes
		# check if it's been more than one minute since the script updated -- it means we woke up from sleep
		$updateDiffInMinutes = $minutesPassed - $minutesPassedLastFrame
		if($updateDiffInMinutes -gt 1) {
			$idleSeconds = 0
			Write-Log "It's been $updateDiffInMinutes minute(s) since the last update, which means we were sleeping or somehow lagging a lot, Resetting idle counter." "INFO"
			if($OnlyThisScriptCanCauseWindowsToSleep -eq $true){
				[WindowsSleepWrangler]::IgnoreIdleTimers()
			}
		}
		$minutesPassedLastFrame = $minutesPassed
		
		# CPU
		$cpu = Get-CpuUsagePercent
		$cpuAbove = $cpu -gt $CpuThresholdPercent
		$cpuHistory.Enqueue($cpuAbove)
		if ($cpuHistory.Count -gt $maxSamples) { 
			$null = $cpuHistory.Dequeue()
		}

		# Disk
		$disk = Get-DiskIoKBps
		$diskAbove = $disk -gt $DiskThresholdKBps
		$diskHistory.Enqueue($diskAbove)
		if ($diskHistory.Count -gt $maxSamples) { 
			$null = $diskHistory.Dequeue()
		}

		# Network, sums up download and upload on all active non-virtual network interfaces
		$net = Get-NetworkIoKBps
		#Write-Log "net: $net" "DEBUG"
		$netAbove = $net -gt $NetworkThresholdKBps
		$netHistory.Enqueue($netAbove)
		if ($netHistory.Count -gt $maxSamples) { 
			$null = $netHistory.Dequeue()
		}
		
		# Audio
		$audioPlaying = Get-AudioIsPlaying
		$audioHistory.Enqueue($audioPlaying)
		if ($audioHistory.Count -gt $maxSamplesAudio) { 
			$null = $audioHistory.Dequeue()
		}

		# Mouse
		$currMouse = Get-MouseMovementPixels
		$mouseDelta = [math]::Sqrt((($currMouse.X - $prevMouse.X) * ($currMouse.X - $prevMouse.X)) + 
									(($currMouse.Y - $prevMouse.Y) * ($currMouse.Y - $prevMouse.Y)))
		$prevMouse = $currMouse
		$mouseMoved = $mouseDelta -gt $MouseThresholdPixels

		# Log every $logToFileIntervalSeconds seconds 
		if ($idleSeconds -ge $logToFileIntervalSeconds) {
			$mouseLog = " "
			if (Test-IsInteractiveSession) {
				$mouseLog = "MouseDelta: $([math]::Round($mouseDelta,1)) px"
			} else {
				$mouseLog = "(mouse check skipped)"
			}
			Write-Log "[idleSeconds: $idleSeconds] CPU: $cpu % | Disk: $disk KBps | Net: $net KBps | $mouseLog | Idle: $([math]::Round($idleSeconds/60,2)) min / $IdleDurationMinutes min" "INFO"
		} elseif ($LogToConsoleVerbose) {
			$mouseLog = " "
			if (Test-IsInteractiveSession) {
				$mouseLog = "MouseDelta: $([math]::Round($mouseDelta,1)) px"
			} else {
				$mouseLog = "(mouse check skipped)"
			}
			Write-Host-Wrapper "[idleSeconds: $idleSeconds] CPU: $cpu % | Disk: $disk KBps | Net: $net KBps | $mouseLog | Idle: $([math]::Round($idleSeconds/60,2)) min / $IdleDurationMinutes min" "INFO"
		}

		#Write-Log "[idleSeconds: $idleSeconds] CPU: $cpu % | Disk: $disk KBps | Net: $net KBps | $mouseLog | Idle: $([math]::Round($idleSeconds/60,2)) min / $IdleDurationMinutes min" "INFO"
		
		# Check for sustained activity (reset idle timer)
		$hasSustainedActivity = $false
		
		# CPU: >= $ActiveSamplesWithinInterval of last $maxSamples were above threshold
		if ($cpuHistory.Count -ge $maxSamples) {
			$activeCount = ($cpuHistory | Where-Object { $_ }).Count
			if ($activeCount -ge $ActiveSamplesWithinInterval) {
				if ($idleSeconds -ge $logToFileIntervalSeconds) {
					Write-Log "[IDLE BREAKER][idleSeconds: $idleSeconds] CPU: $activeCount/$maxSamples samples > $CpuThresholdPercent% (>= $ActiveSamplesWithinInterval required) for $ActivityDetectionPeriod sec" "INFO"
				} elseif ($LogToConsoleVerbose) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $idleSeconds] CPU: $activeCount/$maxSamples samples > $CpuThresholdPercent% (>= $ActiveSamplesWithinInterval required) for $ActivityDetectionPeriod sec" "INFO"
				}
				$hasSustainedActivity = $true
			}
		}

		# Disk: >= $ActiveSamplesWithinInterval of last $maxSamples were above threshold
		if ($diskHistory.Count -ge $maxSamples) {
			$activeCount = ($diskHistory | Where-Object { $_ }).Count
			if ($activeCount -ge $ActiveSamplesWithinInterval) {
				if ($idleSeconds -ge $logToFileIntervalSeconds) {
					Write-Log "[IDLE BREAKER][idleSeconds: $idleSeconds] Disk: $activeCount/$maxSamples samples > $DiskThresholdKBps KBps (>= $ActiveSamplesWithinInterval required) for $ActivityDetectionPeriod sec" "INFO"
				} elseif ($LogToConsoleVerbose) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $idleSeconds] Disk: $activeCount/$maxSamples samples > $DiskThresholdKBps KBps (>= $ActiveSamplesWithinInterval required) for $ActivityDetectionPeriod sec" "INFO"
				}
				$hasSustainedActivity = $true
			}
			
		}

		# Network: >= $ActiveSamplesWithinInterval of last $maxSamples were above threshold
		if ($netHistory.Count -ge $maxSamples) {
			$activeCount = ($netHistory | Where-Object { $_ }).Count
			if ($activeCount -ge $ActiveSamplesWithinInterval) {
				if ($idleSeconds -ge $logToFileIntervalSeconds) {
					Write-Log "[IDLE BREAKER][idleSeconds: $idleSeconds] Network: $activeCount/$maxSamples samples > $NetworkThresholdKBps KBps (>= $ActiveSamplesWithinInterval required) for $ActivityDetectionPeriod sec" "INFO"
				} elseif ($LogToConsoleVerbose) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $idleSeconds] Network: $activeCount/$maxSamples samples > $NetworkThresholdKBps KBps (>= $ActiveSamplesWithinInterval required) for $ActivityDetectionPeriod sec" "INFO"
				}
				$hasSustainedActivity = $true
			}
		}

		# Audio: all samples must be true
		if ($audioHistory.Count -eq $maxSamplesAudio -and ($audioHistory | Where-Object { $_ }).Count -eq $maxSamplesAudio) {
			if($idleSeconds -ge $ActivityDetectionPeriodAudio){
				if ($idleSeconds -ge $logToFileIntervalSeconds) {
					Write-Log "[IDLE BREAKER][idleSeconds: $idleSeconds] Sustained audio playing for $ActivityDetectionPeriodAudio sec, Resetting idle counter." "INFO"
				} elseif ($LogToConsoleVerbose) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $idleSeconds] Sustained audio playing for $ActivityDetectionPeriodAudio sec, Resetting idle counter." "INFO"
				}
				$hasSustainedActivity = $true
			}
		}
		
		if ($mouseMoved) {
			#if ($idleSeconds -ge $ActivityDetectionPeriod) {
				if($idleSeconds -ge $logToFileIntervalSeconds){
					Write-Log "[IDLE BREAKER][idleSeconds: $idleSeconds] Mouse moved $([math]::Round($mouseDelta,1)) px > $MouseThresholdPixels, Resetting idle counter." "INFO"
				} elseif ($LogToConsoleVerbose) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $idleSeconds] Mouse moved $([math]::Round($mouseDelta,1)) px > $MouseThresholdPixels, Resetting idle counter." "INFO"
				}
				$hasSustainedActivity = $true
			#}
		}

		if ($hasSustainedActivity) {
			$idleSeconds = 0
			$cpuHistory.Clear()
			$diskHistory.Clear()
			$netHistory.Clear()
			$audioHistory.Clear()
		} else {
			$idleSeconds += $SampleIntervalSec
		}

		# Poll power plan timeout every $powerPlanPollIntervalSeconds seconds
		$nextPowerPlanPoll--
		if ($nextPowerPlanPoll -le 0) {
			if ($FollowTheSameSleepTimeSettingAsYourPowerPlan) {
				$newTimeout = Get-PowerPlanIdleTimeoutMinutes
				if ($newTimeout -ne $IdleDurationMinutes) {
					Write-Log "Power plan idle timeout changed: $IdleDurationMinutes to $newTimeout min" "INFO"
					$IdleDurationMinutes = $newTimeout
				}
			}
			$nextPowerPlanPoll = $powerPlanPollIntervalSeconds
		}

		# Check if ready to sleep
		if ($idleSeconds -ge ($IdleDurationMinutes * 60)) {
			$abort = Show-AbortDialog -Seconds $SleepAbortWindowCountdownSeconds

			if ($abort) {
				Write-Log "User aborted sleep." "INFO"
				$idleSeconds = 0
			} else {
				Write-Log "Proceeding to sleep because no abort or non-interactive session..." "INFO"
				#if (Test-Path $PauseFlagPath) {
				#    Remove-Item $PauseFlagPath -Force
				#    Write-Log "Deleted flag file: $PauseFlagPath" "INFO"
				#}
				
				if($OnlyThisScriptCanCauseWindowsToSleep -eq $true){
					[WindowsSleepWrangler]::StopIgnoringIdleTimers()
				}
				
				#Enter-SleepState
				$idleSeconds = 0
				Write-Log "System be woke. Resuming monitoring." "INFO"
				
				if($OnlyThisScriptCanCauseWindowsToSleep -eq $true){
					[WindowsSleepWrangler]::IgnoreIdleTimers()
				}
			}
		}
	}
} finally {
	if($OnlyThisScriptCanCauseWindowsToSleep -eq $true){
		[WindowsSleepWrangler]::StopIgnoringIdleTimers()
	}
}
