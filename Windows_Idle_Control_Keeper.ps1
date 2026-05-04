<#
.SYNOPSIS
	
	MIT License, Copyright (c) 2026 Tudor Berechet [tdbe](https://github.com/tdbe) 
	
.SYNOPSIS

	🕯️W.I.C.K. Handles system Idle using your specific thresholds & conditions. It triggers or prevents Windows sleep, hibernate, display, lock, screen saver, on Your terms. Detects activity on CPU, (non-virtual) Network (internet & LAN), Storage, Input, peak Sound value. Idle-breaking events via frequency & amplitude / time periods. Power Plan aware.
	
.DESCRIPTION

	# Windows-Idle-Control-Keeper

	# Intro:

	[Saul Goodman voice] **Can't get reliable sleep? Feeling like it's out of your control? Well fret not, just run this script and you can start counting those sheep!**

	**WICK - Windows Idle Control Keeper**

	This non-admin script detects Idle activity with your specific thresholds and conditions, and triggers / prevents Windows sleep, hibernate, display, lock, screen saver, on Your terms. Detects activity on CPU, (non-virtual) Network (both internet and LAN), Storage, input, and peak Sound value; counting idle-breaking event frequency and amplitude per time periods, to determine if an Idle timer should continue or be broken. It's Windows Power Plan aware, including display off and screensaver schedule, and maintains windows screen locking.

	I don't usually post my system scripts but it annoyed me that for such a wide need, there was nothing out there but forum threads of people using obscure and partial tools like [DontSleep!.exe](https://www.softwareok.com/?Download=DontSleep) [from 2014](https://www.chip.de/downloads/Don-t-Sleep_42626965.html)

	# Features, Dependencies, Log Example, Notes, Run & Parameters:

	## Features:

	- Does not require administrator permission.
	- Works even if windows is locked. Also works if logged out or never logged in (if you start it at system start via task scheduler).
	- Shows warning / abort window for AbortWindowCountdownSeconds before triggering a sleep or hibernate (if in an interactive session (not locked or logged out)).
	- Dynamically reads (every minute (configurable)) from your currently active windows power plan (plugged in or battery) to check sleep and hibernate times.
	- Also can read from Settings_File_Windows_Idle_Control_Keeper_txt. So you can pause or tweak settings while the script is running (every FileSettingsPollIntervalSeconds).
	- Determines idle by accumulating sustained activity events of configurable frequency and amplitude over certain timeframes, through samples every ~second (with delta time): based on if there's CPU, Network, sorage (without waking sleeping hard disks), audio, and input activity.
	- Can prevent windows from sleeping/hibernate until this script decides it's time.
	- Can set a sleep or hibernate time for longer than 5h (the max that Windows power plan allows for some gormless reason).
	- Allows a blacklist for logical drives e.g. `"L", "A", "N"` - you may have drives that have activity you consider passive and you're okay sleeping on. But also keep in mind the NetworkThresholdKBps setting.
	- Can be paused while running by creating an empty `.ignore_running_Windows_Idle_Control_Keeper_script` flag file.
	- Logs what's going on to Windows' Event Viewer - Applicaton Log (only idle on (after 1m of idle) and idle off). Also logs to file at LogPath, so you know at what time Idle state was broken, by what, and after how much idle time. (or if there were errors) (log cleans itself up to stay less than LogMaxSizeMB)
	- It maintains windows screen locking (also can lock on demand), and display off and screensaver schedule (can be triggered on demand).

	## Dependencies:

	- python (and the `checkIfAudioIsPlaying.py` script for custom volume peak thresholds. It requires `pip install pycaw`.)
	- virtually any .net (C# capability) installed on the system (for `SetThreadExecutionState`)
	- powershell 5.1 (the latest is powershell 7+)
	- You need to check that the paths are correct / to your liking. Set the corresponding "*Path" parameters.
	- No admin requirements

	## Log Example:

	```
	[2026-04-30 00:11:58] [INFO] ~*------- W.I.C.K. started. -------
	[2026-04-30 00:11:58] [INFO]   Log path: C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log
	[2026-04-30 00:11:58] [INFO] Power plan sleep idle timeout: 30 min; sleep=1800 sec
	[2026-04-30 00:11:58] [INFO] Power plan hibernate idle timeout: 60 min; hibernate=3600 sec
	[2026-04-30 00:11:58] [INFO]   Using windows power plan's sleep and hiberante values: 30 min and 60 min. (We check to update this value every: SettingsPollIntervalSeconds: 60 sec.)
	[2026-04-30 00:11:58] [INFO] It's been 171368.983134705 minute(s) since the last update, which means we
	were sleeping or somehow lagging a lot, Resetting idle counter.
	[2026-04-30 00:13:38] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:15:49] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:18:03] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:20:18] [INFO] [IDLE BREAKER] Sustained audio playing for 5 sec. Resetting idle counter.
	[2026-04-30 00:22:27] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:24:35] [INFO] [IDLE BREAKER] Network: 3/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:26:53] [INFO] [IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:29:07] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:31:16] [INFO] [IDLE BREAKER] Sustained audio playing for 5 sec. Resetting idle counter.
	[2026-04-30 00:33:22] [INFO] [IDLE BREAKER] Disk: 4/6 samples > 1250 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:35:41] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:38:00] [INFO] [IDLE BREAKER] Network: 3/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:42:30] [INFO] [IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:44:38] [INFO] [IDLE BREAKER] Mouse/touch/keyboard activity registered 1 seconds ago. Resetting idle counter.
	[2026-04-30 00:46:52] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:49:18] [INFO] [IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:51:39] [INFO] [IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:53:57] [INFO] [IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required) for 6 sec
	[2026-04-30 00:53:59] [INFO] CPU: 1 % | Disk: 8279 KBps | Net: 1201 KBps | Input: 91 s ago | Idle: 0 min | (T Sleep: 30 min | T Hibernate: 60 min | T Display: 0 sec | T ScreenSaver: 0 sec)
	```

	## Notes: 
	
	### Note: 
	
	Tested on Windows 11 LTSC, laptop and PC.

	### Note:
	
	This script cannot and will never: ask for administrator privileges, listen to your sounds, key presses, taps and clicks, connect to the internet or network. It only asks for, reads, and logs, when an event of a certain category happened, does not know what data it had.

	### Note:
	
	I built-in a 60 second failsafe, before which this script won't do anything. So you can't screw yourself over: e.g. set a 1s sleep timeout, or force lock after 1s idle etc. So you have 1 minute to fix it after wake, or even on startup, even if you set it to start at system start via task scheduler.

	### Note: 

	Because it doesn't use admin rights, while PreventAndReplaceWindowsAutoSleep is $true, this script can't check if other processes requested that the system not sleep, e.g. an active remote desktop connection while the PC is otherwise within your idle thresholds; so it won't respect their request. You could fix this by either setting PreventAndReplaceWindowsAutoSleep to $false, or by running this script as admin and uncomment the `ctrl+f:[respectOtherApps]` code blocks.

	### Note:

	I've always nuked Modern Standby from every PC I touched, because we have literally 0.0f low-power hardware and protocol standards, and I don't want constant 100W power draw, and for laptops my battery to run out in 2 hours while "sleeping" with the lid closed (Microsoft is the most infuriating thing in the history of ever). You can have a look at how I printed and fetched the sleep AC/DC settings, and figure out the parsing of anything else if you want. Tests & PRs welcome.

	### Note: 

	I wouldn't be caught dead writing for free 2000 lines of powershell script of all things, so for this I tried out LLMs. I don't consider llm output even remotely reliable, but this is all verified and very re-written by me. For those curious: I used 256k context and: qwen 3 coder next 80b a3b q6, qwen 3.6 35b a3b q8, and qwen 3.6 27b q4, locally. They're "great" (within 5-10% of the scores of the huge frontier models) but simultaneously also completely shit at even such a simple job, and not just because this solution doesn't already exist: ie they picked network and storage checks that take at least 1s to return a value, and were calling them repeatedly in loops per disk and per adapter, resulting in a while loop that runs once every 7-10s.. So the verdict is I had to do all the thinking myself. They only oneshotted the logging, the cpu, the sleep functions, and the .PARAM list. Also the peak audio checking I had to research and write myself in python after many wildly off LLM solutions.

	## Run & Parameters:

	There are a LOT of parameters you can set when calling the script or adding it to Task Scheduler (the ones you skip will have defaults).

	To see all parameters and their description, run this command: 

	```
	get-help "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -detailed
	```

	### Run in a powershell terminal window, examples:

	```
	powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$true -PreventAndReplaceWindowsAutoSleep:$true # other flags -etc. -etc.

	or

	powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$false -UserSpecifiedSleepIdleTimeMinutes:720 -PreventAndReplaceWindowsAutoSleep:$true -LockPcAtThisIdleTimeSeconds:300 -TurnOnScreensaverAtThisIdleTimeSeconds:0 -TurnOffDisplayAtThisIdleTimeSeconds:600 # other flags -etc. -etc.

	```

	### Run in Task Scheduler:

	#### Program/script: 

	```
	powershell.exe
	```

	#### Add arguments (window opens as minimized):

	```
	-NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$true -UserSpecifiedSleepIdleTimeMinutes:30 -PreventAndReplaceWindowsAutoSleep:$true #  other flags -etc. -etc.
	```

	#### Add arguments (no window, runs in background completely hidden):

	```
	-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$true -UserSpecifiedSleepIdleTimeMinutes:30 -PreventAndReplaceWindowsAutoSleep:$true #  other flags -etc. -etc.
	```
	
	#### PS:
	
	Instead of regular params, I added $script:Config that loads from either the params ($PSBoundParameters) or the Settings_File_Windows_Idle_Control_Keeper_txt. So you can pause or tweak settings while the script is running (every FileSettingsPollIntervalSeconds).
	
.PARAMETER CliParamsAlwaysOverwriteSettingsFile
  The entries in the settings file, even if default values, will overwrite the command line parameters unless the command line has the -CliParamsAlwaysOverwriteSettingsFile:$true. The settings file is read every FileSettingsPollIntervalSeconds. (default: $false)
  
.PARAMETER PauseScript
  You can pause and resume the script in real time (every FileSettingsPollIntervalSeconds) using this parameter in the settings file. From the cli it will simply start the script as paused (the settings file is still polled while paused). NOTE: This param gets read even if -CliParamsAlwaysOverwriteSettingsFile:$true. (default: $false)

.PARAMETER PreventAndReplaceWindowsAutoSleep
  Uses `SetThreadExecutionState` to prevent idle-based sleep commands. Note that actively triggering Sleep e.g. via the Start menu, or via an explicit function call (e.g. SetSuspendState) from some active software, or laptop lid close, will STILL cause the PC to sleep! Also, if this is set to $false, you will get windows' event plus also this script's event (e.g. windows turns display off (if set to), and also this script turns display off (if set to) - so, whichever comes first). (default: $true)
  
.PARAMETER FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan
  Read idle timeout values from Windows power plan (sleep, hibernate, display). Can still be overridden by the user defined sleep, hibernate, display off, screensaver time, if set to greater than zero (and whichever comes first will be triggered, unless you PreventAndReplaceWindowsAutoSleep:$true). (default: $true)

.PARAMETER UserSpecifiedSleepIdleTimeMinutes
  Using this you can set more than the weird 5h max (e.g. 720 mins (12 hours)) limit that Windows power plan lets you set.  
  Ignored if FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan == $true (unless the power plan is somehow inaccessible). (default: 30)
  
.PARAMETER UserSpecifiedHibernateIdleTimeMinutes
  Using this you can set more than the weird 5h max (e.g. 720 mins (12 hours)) limit that Windows power plan lets you set.  
  Ignored if FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan == $true (unless the power plan is somehow inaccessible). (default: 30)
  
.PARAMETER TurnOnScreensaverAtThisIdleTimeSeconds
  Zero means it doesn't trigger (or the system screensaver setting is used). (default: 0) 
  
.PARAMETER TurnOffDisplayAtThisIdleTimeSeconds
  If nonzero, this does override the power plan when FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan is true. Zero means it doesn't trigger (or the system power setting is used). (default: 0) 
   
.PARAMETER CpuThresholdPercent
  CPU usage above this resets idle timer. (default: 7)

.PARAMETER DiskThresholdKBps
  Disk I/O (KB/s) above this resets idle timer. (default: 1250)

.PARAMETER NetworkThresholdKBps
  Network I/O (KB/s) above this resets idle timer. (default: 850)

.PARAMETER ActiveSamplesWithinInterval
  How many instances (seconds, but there can be lag) of activity must be detected within the last ActivityDetectionPeriodSeconds seconds for us to consider that activity an idle breaker. (default: 3)

.PARAMETER ActivityDetectionPeriodSeconds
  Seconds window to check for sustained activity. (default: 6)

.PARAMETER ActivityDetectionPeriodSecondsAudio
  Separate timeout for audio - counts if there was constant sound in this last period of seconds with a custom sound peak threshold to e.g. avoid background noise. (default: 5)

.PARAMETER LockPcAtThisIdleTimeSeconds 
  If not zero ((and due to failsafe) actually if greater than FailsafeTimeSeconds) will lock pc at this idle time, which can be earlier than when Windows decides to do it. (default: 0)

.PARAMETER IdleSecondsBeforeWeBroadcastSystemIdleEvent
  How much idle time must pass before we declare the system idle as far as this script is concerned, and send an idle event to the windows Event Viewer's Applicaton Log (regardless of when the dysplay is turned off or screensaver turns on or anything else). (default: 60)

.PARAMETER DiskBlacklistDrives
  Allows blacklist for logical drives e.g. `"L", "A", "N"` - drives that have activity but you consider passive and you're okay sleeping on them.  
  Keep in mind the NetworkThresholdKBps setting.  
  (default: @("E", "F"))
  
.PARAMETER Settings_File_Windows_Idle_Control_Keeper_txt
  These settings can be edited while the script is running and the script will read them every FileSettingsPollIntervalSeconds. (default: "C:\Commands_And_Logs\[Settings_File]_Windows_Idle_Control_Keeper.txt")
  
.PARAMETER PycawAudioCheckerPath
  Full path to the Python script used to detect audio playback with custom threshold.  
  (default: "C:\Commands_And_Logs\Pycaw_check_if_audio_is_playing.py")

.PARAMETER PythonPath
  Full path to the Python executable used to run the audio checker script.
  (default: "$env:USERPROFILE\AppData\Local\Programs\Python\Python312\python.exe")

.PARAMETER PauseFlagPath
  Path to a flag file. If this file exists, the script pauses monitoring and skips sleep.  
  Allows manual pause/resume by creating/deleting the file.
  (default: "C:\Commands_And_Logs\.ignore_running_Windows_Idle_Control_Keeper_script")

.PARAMETER LogPath
  Full path to the log file. (default: "C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log")

.PARAMETER LogMaxAgeDays
  Keep logs this many days. (default: 30)

.PARAMETER LogMaxSizeMB
  Rotate log if larger than this MB. (default: 10)

.PARAMETER LogToFileIntervalSeconds
  To prevent writing to file every second while you're using the PC, it won't log unless it's been idle for this many seconds. (default: 60)
  
.PARAMETER LogToConsoleVerbose
  Whether to log to the console (not log file) as often as there is an event in the constant loop. (default: true)

.PARAMETER SleepAbortWindowCountdownSeconds
  Seconds to show the sleep abort dialog before triggering sleep. (default: 60)

.PARAMETER SampleIntervalSec
  How often to sample system metrics. (default: 1)
  
.PARAMETER SettingsPollIntervalSeconds
  Dynamically reads from your currently active windows power plan (plugged in or battery) to check sleep and also hibernate times. (default: 60)
  
.PARAMETER FileSettingsPollIntervalSeconds
  Dynamically reads parameters from your settings file (should run at same interval as SettingsPollIntervalSeconds, unless you're not using it (0 means not used)). (default: 60)
  
.PARAMETER FailsafeTimeSeconds
  I use a failsafe time in case somebody screws something up / adds something that for example would lock the pc every second. This way if you sleep + wake, or restart the pc, you get e.g. 60 seconds to stop it even if you set it to run hidden on system startup from task schedule. (default: 59)
#>

# Note: this doesn't work unless you run the script as administrator, so I commented it out ctrl+f:[respectOtherApps]
#.PARAMETER RespectOtherAppsSleepExecutionPreventionFlags
#  Respects other apps if/when they do what we ourselves do with the -PreventAndReplaceWindowsAutoSleep flag. (default: $false)

# ───────────────────────────────────────────────────────────────────────────────
# 1. PARAMETER BLOCK (CLI overrides only)
# ───────────────────────────────────────────────────────────────────────────────
param(
	[bool]$PauseScript,
	[bool]$CliParamsAlwaysOverwriteSettingsFile,
    [bool]$PreventAndReplaceWindowsAutoSleep,
    [bool]$FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan,
    [int]$UserSpecifiedSleepIdleTimeMinutes,
    [int]$UserSpecifiedHibernateIdleTimeMinutes,
    [int]$TurnOnScreensaverAtThisIdleTimeSeconds,
    [int]$TurnOffDisplayAtThisIdleTimeSeconds,
    [int]$CpuThresholdPercent,
    [int]$DiskThresholdKBps,
    [int]$NetworkThresholdKBps,
    [int]$ActiveSamplesWithinInterval,
    [int]$ActivityDetectionPeriodSeconds,
    [int]$ActivityDetectionPeriodSecondsAudio,
    [int]$LockPcAtThisIdleTimeSeconds,
    [int]$IdleSecondsBeforeWeBroadcastSystemIdleEvent,
    [string[]]$DiskBlacklistDrives,
    [string]$Settings_File_Windows_Idle_Control_Keeper_txt,
    [string]$PycawAudioCheckerPath,
    [string]$PythonPath,
    [string]$PauseFlagPath,
    [string]$LogPath,
    [int]$LogMaxAgeDays,
    [int]$LogMaxSizeMB,
    [int]$LogToFileIntervalSeconds,
    [bool]$LogToConsoleVerbose,
    [int]$SleepAbortWindowCountdownSeconds,
    [int]$SampleIntervalSec,
    [int]$SettingsPollIntervalSeconds,
    [int]$FileSettingsPollIntervalSeconds,
    [int]$FailsafeTimeSeconds
)

# ───────────────────────────────────────────────────────────────────────────────
# 2. CONFIGURATION DEFAULTS (single source of truth)
# ───────────────────────────────────────────────────────────────────────────────
$script:Config = @{
	PauseScript 											= $false
	CliParamsAlwaysOverwriteSettingsFile					= $false
    PreventAndReplaceWindowsAutoSleep           			= $true
    FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan 	= $true
    UserSpecifiedSleepIdleTimeMinutes           			= 30
    UserSpecifiedHibernateIdleTimeMinutes       			= 30
    TurnOnScreensaverAtThisIdleTimeSeconds      			= 0
    TurnOffDisplayAtThisIdleTimeSeconds         			= 0
    CpuThresholdPercent                         			= 7
    DiskThresholdKBps                           			= 1250
    NetworkThresholdKBps                        			= 850
    ActiveSamplesWithinInterval                 			= 3
    ActivityDetectionPeriodSeconds              			= 6
    ActivityDetectionPeriodSecondsAudio         			= 5
    LockPcAtThisIdleTimeSeconds                 			= 0
    IdleSecondsBeforeWeBroadcastSystemIdleEvent 			= 60
    DiskBlacklistDrives                         			= @("E", "F")
    Settings_File_Windows_Idle_Control_Keeper_txt 			= "C:\Commands_And_Logs\[Settings_File]_Windows_Idle_Control_Keeper.txt"
    PycawAudioCheckerPath                       			= "C:\Commands_And_Logs\Pycaw_check_if_audio_is_playing.py"
    PythonPath                                  			= "$env:USERPROFILE\AppData\Local\Programs\Python\Python312\python.exe"
    PauseFlagPath                               			= "C:\Commands_And_Logs\.ignore_running_Windows_Idle_Control_Keeper_script"
    LogPath                                     			= "C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log"
    LogMaxAgeDays                               			= 30
    LogMaxSizeMB                                			= 10
    LogToFileIntervalSeconds                    			= 60
    LogToConsoleVerbose                         			= $true
    SleepAbortWindowCountdownSeconds            			= 60
    SampleIntervalSec                           			= 1
    SettingsPollIntervalSeconds                 			= 60
    FileSettingsPollIntervalSeconds                 		= 60
    FailsafeTimeSeconds                         			= 59
}

# ───────────────────────────────────────────────────────────────────────────────
# 3. INITIALIZE CONFIG FROM CLI PARAMETERS (if provided)
# ───────────────────────────────────────────────────────────────────────────────
$PSBoundParametersCount = $PSBoundParameters.Count
if ($PSBoundParametersCount -gt 0) {
	Write-Warning "~~~~ Loading CLI Params: PSBoundParametersCount: $PSBoundParametersCount."
    foreach ($key in $PSBoundParameters.Keys) {
		$configContainsKey = $script:Config.ContainsKey($key)
		Write-Warning "~~~~~~~~ script:Config.ContainsKey($key): $configContainsKey."
		if ($configContainsKey -eq $true) {
			$configAtKey = $script:Config[$key]
			Write-Warning "~~~~~~~~~~~~ default script:Config[key]: $configAtKey."
			$PSBoundParametersAtKey = $PSBoundParameters[$key]
			Write-Warning "~~~~~~~~~~~~ CLI PSBoundParameters[key]: $PSBoundParametersAtKey."
			if(-not ($script:Config[$key] -eq $PSBoundParameters[$key])) {
				Write-Warning ">>~~~~~~~~~~~~>> Applying the non-default value coming from the command line parameter."
				$script:Config[$key] = $PSBoundParameters[$key]
			}
        } else {
            Write-Error "[!]~~~~~~~ Parameter '$key' not recognized ignored."
        }
    }
}

# ───────────────────────────────────────────────────────────────────────────────
# 4. SETTINGS FILE PARSER (supports comments, quoted values, arrays, booleans)
# ───────────────────────────────────────────────────────────────────────────────
function Expand-EnvironmentVariables {
    <#
    .SYNOPSIS
        Safely expands $env:VAR references in strings read from config files.
    #>
    param([string]$InputString)
    # Convert $env:VAR to %VAR% so .NET can expand it
    $windowsStyle = $InputString -replace '\$env:(\w+)', '%$1%'
    try {
        return [Environment]::ExpandEnvironmentVariables($windowsStyle)
    }
    catch {
        # If a variable doesn't exist, return the original string unchanged
        return $InputString
    }
}
function Update-ConfigFromSettingsFile {
	
    $Path = $script:Config['Settings_File_Windows_Idle_Control_Keeper_txt']
	
    if (-not (Test-Path -LiteralPath $Path)) { 
		Write-Warning "~~~~ Settings file not found at path: $Path."
		return 
	}
	
	Write-Host "||||~~~~ Updating config settings from file: $Path"
    $content = Get-Content -LiteralPath $Path
	
    $lines = $content -split '\r?\n' | ForEach-Object {
        # Strip comments (# outside quotes)
        $i = 0; $inQuote = $false; $quote = $null
        while ($i -lt $_.Length) {
            $c = $_[$i]
            if ($c -eq '"' -or $c -eq "'") {
                if (-not $inQuote) { $inQuote = $true; $quote = $c }
                elseif ($c -eq $quote) { $inQuote = $false; $quote = $null }
            }
            elseif ($c -eq '#' -and -not $inQuote) {
                $_ = $_.Substring(0, $i); break
            }
            $i++
        }
        $_.Trim()
    } | Where-Object { $_ -and $_ -notmatch '^#' }

    foreach ($line in $lines) {
        if ($line -match '^(\w+):\s*(.*)$') {
            $key = $matches[1]
            $valStr = $matches[2].Trim()
            if ($script:Config.ContainsKey($key)) {
                try {
                    $val = if ($valStr -eq '$true') { $true }
                           elseif ($valStr -eq '$false') { $false }
                           elseif ($valStr -match '^@\(.*\)$') {
                               # Parse @("a", "b") → @("a","b")
                               $inner = $valStr.Substring(2, $valStr.Length - 3)
                               $items = @()
                               $current = ''; $inQ = $false; $q = $null
                               foreach ($c in $inner.ToCharArray()) {
                                   if ($c -eq '"' -or $c -eq "'") {
                                       if (-not $inQ) { $inQ = $true; $q = $c }
                                       elseif ($c -eq $q) { $inQ = $false; $q = $null }
                                       $current += $c
                                   }
                                   elseif ($c -eq ',' -and -not $inQ) {
                                       $items += $current.Trim().Trim('"').Trim("'")
                                       $current = ''
                                   }
                                   else { $current += $c }
                               }
                               if ($current) { $items += $current.Trim().Trim('"').Trim("'") }
                               ,$items
                           }
                           elseif ([int]::TryParse($valStr, [ref]0)) { [int]$valStr }
                           else { 
							   # Strip surrounding quotes if present
                               $cleaned = $valStr.Trim('"').Trim("'")
							   # Expand environment variables (e.g., $env:USERPROFILE)
                               Expand-EnvironmentVariables $cleaned
						   }
					
					if ($script:Config['CliParamsAlwaysOverwriteSettingsFile'] -eq $true) {
						Write-Host "Skipping loading: $key ($val) from file because it's set to be overwritten by the CLI param version: $cliVersion"
					} 
					
					if ($script:Config[$key] -ne $val -and $script:Config['CliParamsAlwaysOverwriteSettingsFile'] -eq $false -or $key -eq 'PauseScript') {
						Write-Host "||||>>~~~~>> Updated config from file: '$key': from $($script:Config[$key]) to $val"
						$script:Config[$key] = $val
					}
                }
                catch {
                    Write-Warning "Failed to parse '$valStr' for key '$key': $_"
                }
            }
        }
    }
}
# ───────────────────────────────────────────────────────────────────────────────
# ^ AI boilerplate to read from settings file as well as from cli params, finished. Real script starts now:
# ───────────────────────────────────────────────────────────────────────────────

#[int]$script:Config['SampleIntervalSec'] = 1
# Note: this doesn't work unless you run the script as administrator, so I commented it out ctrl+f:[respectOtherApps]
#[bool]$RespectOtherAppsSleepExecutionPreventionFlags = $false, # Respects other apps if/when they do what we ourselves do with the -PreventAndReplaceWindowsAutoSleep flag.

$script:g_CurrentSleepIdleTimeMinutes = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
$script:g_CurrentHibernateIdleTimeMinutes = $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
$script:g_DisplayTimeoutDurationSeconds = 0
$script:g_DisplayTurnedOff = $false
$script:g_ScreensaverTimeoutDurationSeconds = 0
$script:g_ScreenSaverStarted = $false
$script:g_PcLockedOnDemand = $false
$script:g_PreventSleep_ES = $false

$script:g_myUnixTimeEpochStart = Get-Date '2026-01-01'
$script:g_minutesPassedLastFrame = 0


# --- Logging Setup ---
# logs to Windows > Event Viewer > Windows Logs > Applicaton. It will have the date and time of the event. These can be queried by scripts.
function LogSystemEvent_IdleOn {
    [CmdletBinding()]
    param()
	
	$script:g_isIdle = $true

    $LogName   = "Application" # writing to the "System" log requires admin privileges
    $Source    = "Application" # writing to your new custom source e.g. "wick_idle_on" requires admin privileges
	$Category  = 69
    $EventId   = 420
    $EntryType = [System.Diagnostics.EventLogEntryType]::Warning # or Information
    $Message   = "[WICK: IDLE] System is idle according to the Windows_Idle_Control_Keeper.ps1. (IdleSecondsBeforeWeBroadcastSystemIdleEvent: $script:Config['IdleSecondsBeforeWeBroadcastSystemIdleEvent'].)"

    Write-EventLog -LogName $LogName -Source $Source -EventId $EventId -EntryType $EntryType -Message $Message -Category $Category
}

# logs to Windows > Event Viewer > Windows Logs > Applicaton. It will have the date and time of the event. These can be queried by scripts.
function LogSystemEvent_IdleOff {
    [CmdletBinding()]
    param()

	$script:g_isIdle = $false

    $LogName   = "Application" # writing to the "System" log requires admin privileges
    $Source    = "Application" # writing to your new custom source e.g. "wick_idle_on" requires admin privileges
    $Category  = 69
    $EventId   = 420
    $EntryType = [System.Diagnostics.EventLogEntryType]::Warning # or Information
    $Message   = "[WICK: NOT Idle] System stopped being idle according to the Windows_Idle_Control_Keeper.ps1"

    Write-EventLog -LogName $LogName -Source $Source -EventId $EventId -EntryType $EntryType -Message $Message -Category $Category
}

$script:g_LogDir = Split-Path $script:Config['LogPath'] -Parent
if (-not (Test-Path $script:g_LogDir)) {
    try {
        New-Item -Path $script:g_LogDir -ItemType Directory -Force | Out-Null
        Write-Host-Wrapper "Created log directory: $script:g_LogDir" "INFO"
    }
    catch {
        Write-Host-Wrapper "ERROR: Could not create log directory: $script:g_LogDir. Logging disabled." "ERROR"
        $script:Config['LogPath'] = $null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    if (-not $script:Config['LogPath']) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

	if($logEntry -eq "WARN") {
		Write-Warning "$logEntry"
	} elseif ($logEntry -eq "ERROR") {
		Write-Error "$logEntry"
	} else {
		Write-Host "[(logged)] $logEntry"
	}

    try {
        Add-Content -Path $script:Config['LogPath'] -Value $logEntry -ErrorAction Stop

        $file = Get-Item $script:Config['LogPath'] -ErrorAction SilentlyContinue
        if ($file -and $file.Length -gt ($script:Config['LogMaxSizeMB'] * 1MB)) {
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
    param([string]$Path = $script:Config['LogPath'])
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

    Get-ChildItem $script:g_LogDir -Filter "IdleSleepMonitor.log.*" -File |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$script:Config['LogMaxAgeDays']) } |
        Remove-Item -Force
}
# --- /Logging Setup ---

#always add type definitions outside of functions (and do the null check) otherwise you're compiling code every function call and also potentially leaking
$script:g_typeName = 'WindowsSleepWrangler'
if (-not ($script:g_typeName -as [type])) {
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
				// NOTE: ON LAPTOPS, ES_SYSTEM_REQUIRED, WILL "INCLUDE"  ES_DISPLAY_REQUIRED (check via `powercfg /requests`), so you need to turn off the display manually instead of relying on windows doing it.
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


# powercfg /requests
#SYSTEM:
#[DRIVER] Realtek High Definition Audio (INTELAUDIO\FUNC_01&VEN_10EC&DEV_0285&SUBSYS_10431493&REV_1000\5&1f865b95&0&0001)
#An audio stream is currently in use.
#[PROCESS] \Device\HarddiskVolume3\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

# check if other processes reqeusted to not idle sleep
# Note: this doesn't work unless you run the script as administrator, so I commented it out ctrl+f:[respectOtherApps]
# but, we can't run powercfg /requests without admin so we can't use this actually..
function Test-OtherSystemExecutionStateHeld {
    [CmdletBinding()]
    param(
        [bool]$areWePreventingIdle = $true
    )

    $output = powercfg /requests 2>&1
    #if ($LASTEXITCODE -ne 0) { return $false }

    $lines = $output -split "`r?`n"
    $inSystemBlock = $false
    $systemEntries = @()

    foreach ($line in $lines) {
        $trimmed = $line.Trim()

        if ($trimmed -eq "SYSTEM:") {
            $inSystemBlock = $true
            continue
        }
		
		if ($trimmed -eq "AWAYMODE:") {
            $inSystemBlock = $false
            break
        }

        if ($inSystemBlock) {
            # Stop block parsing at empty line or next section header
            #if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed -match #"^(AWAYMODE|DISPLAY|SERVICE|DRIVER|GLOBAL|APPLICATION):") {
            #    break
            #}
            
            # Count lines that start with [*] pattern
            if ($trimmed -match '^\s*\[[^\]]+\]\s') {
                $systemEntries += $line
            }
        }
    }

    $count = $systemEntries.Count

    if ($count -ge 2) {
        return $true
    }

    if ($count -eq 1) {
        $isOurPS = $systemEntries[0] -match 'powershell\.exe'
        if ($isOurPS -and $areWePreventingIdle) {
            return $true
        }
        return $false
    }

    return $false
}

# plugged in means AC power in power plan, battery means DC power. Important for sleep timers (different per AC / DC)
function IsComputerPluggedIn{
	return (Get-WmiObject -Class BatteryStatus -Namespace root\wmi).PowerOnLine
}

#always add type definitions outside of functions (and do the null check) otherwise you're compiling code every function call and also potentially leaking
$script:g_typeName = 'Display'
if (-not ($script:g_typeName -as [type])) {
	Add-Type -TypeDefinition @"
	using System;
	using System.Runtime.InteropServices;
	public static class Display {
		[DllImport("user32.dll")]
		private static extern IntPtr SendMessageTimeout(
			IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam,
			int fuFlags, int uTimeout, out IntPtr lpdwResult);

		private const int WM_SYSCOMMAND = 0x0112;
		private const int SC_MONITORPOWER = 0xF170;
		private const int MONITOR_OFF = 2;

		public static void TurnOff() {
			IntPtr result;
			SendMessageTimeout(
				new IntPtr(0xFFFF), // HWND_BROADCAST
				WM_SYSCOMMAND,
				(IntPtr)SC_MONITORPOWER,
				(IntPtr)MONITOR_OFF,
				0, 5000, out result);
		}
	}
"@ -Language CSharp
}

function Turn-Display-Off {
	# it's possible that if thread execution state is set to ES_SYSTEM_REQUIRED (to prevent auto sleep), windows will also not lock the desktop - the security & locking side of things is obscure and may vary by version or policy or drivers even.
	if ($script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true) {
		Lock-PC
	}
	
	$script:g_DisplayTurnedOff = $true
	Write-Log "Turning off Display. g_DisplayTurnedOff: $script:g_DisplayTurnedOff" "Info"
	[Display]::TurnOff()
}

#always add type definitions outside of functions (and do the null check) otherwise you're compiling code every function call and also potentially leaking
$script:g_typeName = 'SystemApi'
if (-not ($script:g_typeName -as [type])) {
	Add-Type -TypeDefinition @'
		using System;
		using System.Runtime.InteropServices;
		public static class SystemApi {
			[DllImport("user32.dll")]
			public static extern bool LockWorkStation();
		}
'@ -Language CSharp
}

function Lock-PC {	
	$script:g_PcLockedOnDemand = $true
	$result = [SystemApi]::LockWorkStation()
	if (-not $result) {
		$script:g_PcLockedOnDemand = $false
		$err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
		Write-Log "Failed to lock workstation." "ERROR"
	} else {
		Write-Log "Locked PC. g_PcLockedOnDemand: $script:g_PcLockedOnDemand" "INFO"
	}
}

#always add type definitions outside of functions (and do the null check) otherwise you're compiling code every function call and also potentially leaking
$script:g_typeName = 'Screensaver'
if (-not ($script:g_typeName -as [type])) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public static class Screensaver {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(
            IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_SYSCOMMAND = 0x0112;
        private const uint SC_SCREENSAVE = 0xF140;

        public static void Trigger() {
            // Broadcast to all top-level windows
            PostMessage((IntPtr)0xFFFF, WM_SYSCOMMAND, (IntPtr)SC_SCREENSAVE, IntPtr.Zero);
        }
    }
"@ -Language CSharp
}

function Start-Screensaver {
	Write-Log "Starting Screen Saver (if it exists)." "Info"
	$script:g_ScreenSaverStarted = $true
    [Screensaver]::Trigger()
}

function Get-ScreensaverTimeoutSeconds {
    # Registry path for screen saver settings (per-user)
    $regPath = "HKCU:\Control Panel\Desktop"
    
    # Check if key exists
    if (-not (Test-Path $regPath)) {
        Write-Log "Registry path '$regPath' not found." "WARN"
        return $null
    }

    # Get values
    $ssActive   = (Get-ItemProperty $regPath -Name "ScreenSaveActive" -ErrorAction SilentlyContinue).ScreenSaveActive
    $ssTimeout  = (Get-ItemProperty $regPath -Name "ScreenSaveTimeOut" -ErrorAction SilentlyContinue).ScreenSaveTimeOut
    $ssSecure   = (Get-ItemProperty $regPath -Name "ScreenSaverIsSecure" -ErrorAction SilentlyContinue).ScreenSaverIsSecure

	# Write-Log "~~~~~~~~~~Screensaver: ssActive: $ssActive, ssTimeout: $ssTimeout, ssSecure: $ssSecure" "WARN"

    # Validate
    if ($ssActive -ne "1") {
        Write-Log "Screensaver is disabled (ScreenSaveActive ≠ 1)." "WARN"
        return $null
    }

    # Convert timeout to seconds (string → int)
    if ([int]::TryParse($ssTimeout, [ref]0)) {
        return [int]$ssTimeout
    } else {
        Write-Log "Invalid ScreenSaverTimeout value: '$ssTimeout'" "WARN"
        return $null
    }
}

function Get-DisplayTimeoutSeconds {
    [CmdletBinding()]
    param(
        [bool]$isAC = $false  # $true = AC, $false = DC
    )

    # Get active plan GUID
    $activePlanLine = powercfg /getactivescheme
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Failed to get active power scheme. Exit code: $LASTEXITCODE" "ERROR"
        return $null
    }
    $activePlan = ($activePlanLine -replace '^.*\(([0-9a-f-]+)\)$', '$1')
    if (-not $activePlan) {
        Write-Log "Failed to parse power plan GUID." "ERROR"
        return $null
    }
	
	$activePlan = $activePlan.Substring(19, 36).Trim()

    # GUIDs
    $SUB_DISPLAY = "7516b95f-f776-4464-8c53-06167f40cc99"

    # Query subgroup only (powercfg shows both AC/DC in one output)
	#$output = powercfg /q $activePlan 2>&1
    $output = powercfg /q $activePlan $SUB_DISPLAY 2>&1
	#Write-Log "~~~~~~ activePlan: $activePlan, SUB_DISPLAY: $SUB_DISPLAY, output: $output" "WARN"
    if ($LASTEXITCODE -ne 0) {
        Write-Log "powercfg failed for display: $output" "ERROR"
        return $null
    }

	$acDcString = "DC"
	if($isAC) {
		$acDcString = "AC"
	}
    $pattern = "Current ${acDcString} Power Setting Index:\s+0x([0-9a-fA-F]+)"
    $line = $output | Select-String -Pattern $pattern -CaseSensitive | Select-Object -First 1

    if ($line) {
        $value = $line.Line -replace ".*Current ${acDcString} Power Setting Index:\s+0x([0-9a-fA-F]+).*", '$1'
		# value is of this format: 00000078
		$decimalValue = 0
        if ([int]::TryParse($value, [System.Globalization.NumberStyles]::HexNumber, $null, [ref]$decimalValue)) {
			return $decimalValue
		} else {
			Write-Log "Failed to parse hex value: '$value'" "WARN"
			return $null
		}
    }

    Write-Log "Could not find 'Current $acDcString Power Setting Index' in powercfg output." "ERROR"
    return $null
}


# --- Helper: Read idle timeout from power plan (in minutes) ---
function Get-PowerPlanIdleTimeoutMinutes {
	[CmdletBinding()]
    [OutputType([PSCustomObject])]
	param(
        [bool]$isAC = $false
    )
    try {
        # Get active power plan GUID
		$activePlan = (powercfg /getactivescheme) -replace '^([0-9a-f-]+).*', '$1'
        if (-not $activePlan) { throw "Failed to get active power plan" }
		
		$activePlan = $activePlan.Substring(19, 36).Trim()
		
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
		
        # Convert to minutes, fallback if disabled
        if ($sleepSec -le 0) {
            Write-Log "Power plan sleep idle timeout: disabled, using fallback: $script:Config['UserSpecifiedSleepIdleTimeMinutes'] min" "INFO"
			Write-Log "sleepSec: $script:Config['UserSpecifiedSleepIdleTimeMinutes']" "INFO"
            return $script:Config['UserSpecifiedSleepIdleTimeMinutes']
        }
        $sleepMinutes = [math]::Round($sleepSec / 60)
        if($script:Config['UserSpecifiedSleepIdleTimeMinutes'] -ne $sleepMinutes) {
			Write-Log "Power plan sleep idle timeout: $sleepMinutes min; sleep=$sleepSec sec" "INFO"
		}
		
		# Convert to minutes, fallback if disabled
        if ($hibernateSec -le 0) {
            Write-Log "Power plan hibernate idle timeout: disabled, using fallback: $script:Config['UserSpecifiedHibernateIdleTimeMinutes'] min" "INFO"
			Write-Log "sleepSec: $script:Config['UserSpecifiedHibernateIdleTimeMinutes']" "INFO"
            return $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
        }
        $hibernateMinutes = [math]::Round($hibernateSec / 60)
        if($script:Config['UserSpecifiedHibernateIdleTimeMinutes'] -ne $hibernateMinutes) {
			Write-Log "Power plan hibernate idle timeout: $hibernateMinutes min; hibernate=$hibernateSec sec" "INFO"
		}
				
		# Return named properties (PS 5.2 compatible)
		[PSCustomObject]@{
			sleepMinutesVal = $sleepMinutes
			hibernateMinutesVal = $hibernateMinutes
		}
    }
    catch {
        Write-Log "Failed to read power plan: $_, using fallback sleep: $script:Config['UserSpecifiedSleepIdleTimeMinutes'] min, and fallback hibernate: $script:Config['UserSpecifiedHibernateIdleTimeMinutes'] min" "WARN"
		# Return named properties (PS 5.2 compatible)
		[PSCustomObject]@{
			sleepMinutesVal = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
			hibernateMinutesVal = $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
		}
    }
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

#always add type definitions outside of functions (and do the null check) otherwise you're compiling code every function call and also potentially leaking
$script:g_typeName = 'PowerManagement'
if (-not ($script:g_typeName -as [type])) {
	Add-Type -TypeDefinition @'
    using System;
    using System.Runtime.InteropServices;
    public static class PowerManagement {
        [DllImport("powrprof.dll", SetLastError = true)]
        public static extern bool SetSuspendState(bool hiberate, bool forceCritical, bool disableWakeEvent);
    }
'@ -Language CSharp
}

# --- Sleep API ---
function Enter-SleepState {
    $result = [PowerManagement]::SetSuspendState($false, $true, $false)
    if (-not $result) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Log "Sleep failed! Win32 error: $err" "ERROR"
    }
}

function Enter-HibernateState {
    $result = [PowerManagement]::SetSuspendState($true, $true, $false)
    if (-not $result) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Log "Hibernate failed! Win32 error: $err" "ERROR"
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
            if ($script:Config['DiskBlacklistDrives'] -contains $driveLetter) { return }
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

#always add type definitions outside of functions (and do the null check) otherwise you're compiling code every function call and also potentially leaking
$script:g_typeName = 'InputMonitor'
if (-not ($script:g_typeName -as [type])) {
	Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    public static class InputMonitor {
        [StructLayout(LayoutKind.Sequential)]
        public struct LASTINPUTINFO {
            public uint cbSize;
            public uint dwTime;
        }

        [DllImport("user32.dll")]
        public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [DllImport("kernel32.dll")]
        public static extern uint GetTickCount();
    }
"@ -Language CSharp
}

#this is better than the `Get-MouseMovementPixels` because it tests for any key events or touchscreen in addition to mouse
#GetLastInputInfo tells you exactly when the last keyboard/mouse/touch event happened
function Get-SecondsSinceLastInputInfo {
	$retIsInteractiveSession = Test-IsInteractiveSession
    if (-not $retIsInteractiveSession) {
        #Write-Log "Mouse check skipped because non-interactive session." "DEBUG"
        return $null
    }
	
    $lii = New-Object InputMonitor+LASTINPUTINFO
    $lii.cbSize = [int][System.Runtime.InteropServices.Marshal]::SizeOf($lii)
    
    if ([InputMonitor]::GetLastInputInfo([ref]$lii)) {
        $idleMs = [InputMonitor]::GetTickCount() - $lii.dwTime
        return [int]($idleMs / 1000)
    }
    return $null
}

# --- Mouse movement check (interactive only) ---
function Get-MouseMovementPixels {
	$retIsInteractiveSession = Test-IsInteractiveSession
    if (-not $retIsInteractiveSession) {
        #Write-Log "Mouse check skipped because non-interactive session." "DEBUG"
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
	$pythonExe = $script:Config['PythonPath']
	$pyScript  = $script:Config['PycawAudioCheckerPath']
	$rawOutput = & "$pythonExe" "$pyScript" 2>$null | ForEach-Object { $_.Trim().ToLower() }
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
    $msg = "Your PC is idle for $script:g_CurrentSleepIdleTimeMinutes minutes.`n`n" +
           "It will sleep in $Seconds seconds.`n" +
           "Press OK to abort."

    $wsh = New-Object -ComObject WScript.Shell
    $result = $wsh.Popup($msg, $Seconds, $title, 4 + 32)

    return ($result -eq 1)
}

# --- Main Loop ---
Write-Log "~*------- W.I.C.K. started. -------"
Write-Log "  Log path: $($script:Config['LogPath'])"

$script:g_isPluggedIn = IsComputerPluggedIn

#Write-Log "  Dynamic idle timeout (checking the current active power plan value every: $($script:Config['SettingsPollIntervalSeconds']))"
if ($script:Config['FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan']) {
	$tuple = Get-PowerPlanIdleTimeoutMinutes $script:g_isPluggedIn
	$script:g_CurrentSleepIdleTimeMinutes = $tuple.sleepMinutesVal
	$script:g_CurrentHibernateIdleTimeMinutes = $tuple.hibernateMinutesVal
    Write-Log "  Using windows power plan's sleep and hiberante values: $($script:g_CurrentSleepIdleTimeMinutes) min and $($script:g_CurrentHibernateIdleTimeMinutes) min. (We check to update this value every: SettingsPollIntervalSeconds: $($script:Config['SettingsPollIntervalSeconds']) sec.)" "INFO"
} else {
	$script:g_CurrentSleepIdleTimeMinutes = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
    Write-Log "  Using manual idle timeout value: system considered idle at: $($script:Config['UserSpecifiedSleepIdleTimeMinutes']) min." "INFO"
}

if ($script:Config['PreventAndReplaceWindowsAutoSleep']) {
	Write-Log "Because PreventAndReplaceWindowsAutoSleep: $($script:Config['PreventAndReplaceWindowsAutoSleep']), we need to manually trigger the display to turn off at the power plan's display setting time (and also lock), and also trigger the screensaver at its time if it exists:" "INFO"
		
	if ($script:Config['TurnOffDisplayAtThisIdleTimeSeconds'] -gt 0) {
		$script:g_DisplayTimeoutDurationSeconds = $script:Config['TurnOffDisplayAtThisIdleTimeSeconds']
		Write-Log "  Using user set display timeout: $($script:Config['TurnOffDisplayAtThisIdleTimeSeconds'])" "INFO"
	} else {
		$script:g_DisplayTimeoutDurationSeconds = Get-DisplayTimeoutSeconds $script:g_isPluggedIn
		Write-Log "  Using windows power plan's g_DisplayTimeoutDurationSeconds: $($script:g_DisplayTimeoutDurationSeconds). We need this because Laptops will stop auto turning off their display if you tell them to not sleep (by using SetThreadExecutionState ES_SYSTEM_REQUIRED). (We check to update this value every: SettingsPollIntervalSeconds: $($script:Config['SettingsPollIntervalSeconds']) sec.)" "INFO"
	}
	
	if($script:Config['TurnOnScreensaverAtThisIdleTimeSeconds'] -gt 0) {
		$script:g_ScreensaverTimeoutDurationSeconds = $script:Config['TurnOnScreensaverAtThisIdleTimeSeconds']
		Write-Log "  Using user set screensaver timeout: $($script:Config['TurnOnScreensaverAtThisIdleTimeSeconds'])" "INFO"
	} else {
		$script:g_ScreensaverTimeoutDurationSeconds = Get-ScreensaverTimeoutSeconds
		Write-Log "  Using windows power plan's g_ScreensaverTimeoutDurationSeconds: $($script:g_ScreensaverTimeoutDurationSeconds) (in case you use the screensaver). We need this because Laptops will stop auto turning on the screensaver if you tell them to not sleep (by using SetThreadExecutionState ES_SYSTEM_REQUIRED). (We check to update this value every: SettingsPollIntervalSeconds: $($script:Config['SettingsPollIntervalSeconds']) sec.)" "INFO"
	}
	
} elseif ($FollowTheSameSleepTimeSettingAsYourPowerPlan -eq $false) {
	if($script:Config['TurnOnScreensaverAtThisIdleTimeSeconds'] -gt 0) {
		$script:g_ScreensaverTimeoutDurationSeconds = $script:Config['TurnOnScreensaverAtThisIdleTimeSeconds']
	}
	if ($script:Config['TurnOffDisplayAtThisIdleTimeSeconds'] -gt 0) {
		$script:g_DisplayTimeoutDurationSeconds = $script:Config['TurnOffDisplayAtThisIdleTimeSeconds']
	}
}

# Load WinForms once (for mouse check)
#try {
#    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
#}
#catch {
#    Write-Log "Failed to load WinForms: $_" "WARN"
#}
# Get initial mouse position, replaced with Get-SecondsSinceLastInputInfo
#$prevMouse = Get-MouseMovementPixels

# Sliding windows for sustained detection
$script:g_maxSamples = [int]([math]::Ceiling($script:Config['ActivityDetectionPeriodSeconds']))# / $script:Config['SampleIntervalSec']))
$script:g_maxSamplesAudio = [int]([math]::Ceiling($script:Config['ActivityDetectionPeriodSecondsAudio']))# / $script:Config['SampleIntervalSec']))
$script:g_idleSeconds = 0.0
$script:g_isIdle = $false
$script:g_sw = [System.Diagnostics.Stopwatch]::StartNew()
$script:g_lastElapsedSeconds = 0.0

$script:g_cpuHistory = New-Object 'System.Collections.Queue' $script:g_maxSamples
$script:g_diskHistory = New-Object 'System.Collections.Queue' $script:g_maxSamples
$script:g_netHistory = New-Object 'System.Collections.Queue' $script:g_maxSamples
$script:g_audioHistory = New-Object 'System.Collections.Queue' $script:g_maxSamplesAudio

$script:g_nextSettingsPoll = $script:Config['SettingsPollIntervalSeconds']
$script:g_nextFileSettingsPoll = $script:Config['FileSettingsPollIntervalSeconds']

if (Test-Path $script:Config['PauseFlagPath']) {
	Write-Log "Script pause flag file is present - while loop running but skipping until flag removed or renamed." "INFO"
}

# NOTE: if this script is frozen or closed unexpectedly in a way that [WindowsSleepWrangler]::StopIgnoringIdleTimers() doesn't get called, windows may not go to sleep again until it is called or until it's restarted. But since we use try - finally, it should auto clean up after itself unless you somehow freeze the thread.
if($script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true){
	$script:g_PreventSleep_ES = $true
	[WindowsSleepWrangler]::IgnoreIdleTimers()
}

try {
	while ($true) {
		#Write-Log "Tick. idleSeconds: $script:g_idleSeconds" "INFO"
		#Write-Host-Wrapper "Tick. idleSeconds: $script:g_idleSeconds" "INFO"
		
		$currentElapsedSeconds = $script:g_sw.Elapsed.TotalSeconds
		$deltaTimeSeconds = $currentElapsedSeconds - $script:g_lastElapsedSeconds
		$script:g_lastElapsedSeconds = $currentElapsedSeconds
		
		# Using delta time we account for any lag spikes. Also we account for long spike from an eventual system sleep and resume.
		Start-Sleep -Milliseconds ([int]([math]::Max(0, 1000-$deltaTimeSeconds*1000)))
		
		$minutesPassed = (Get-Date).Subtract($script:g_myUnixTimeEpochStart).TotalMinutes
		# check if it's been more than one minute since the script updated -- it means we woke up from sleep
		$updateDiffInMinutes = $minutesPassed - $script:g_minutesPassedLastFrame
		if($updateDiffInMinutes -gt 1) {
			$script:g_idleSeconds = -1.0 * $script:Config['FailsafeTimeSeconds'] # we should set it to 0 here but I do a failsafe time here in case somebody screws something up / adds something that for example would lock the pc every second. This way if you sleep + wake, or restart the pc, you get 60 seconds to stop it even if you set it to run hidden on system startup from task scheduler.
			
			$deltaTimeSeconds = 0
			Write-Log "It's been $updateDiffInMinutes minute(s) since the last update, which means we just started, or were sleeping, or somehow lagging a lot. Resetting idle counter, with failsafe, to: $script:g_idleSeconds." "INFO"
		}
		$script:g_minutesPassedLastFrame = $minutesPassed
		
		if ($script:g_idleSeconds -lt 0) {
			$script:g_idleSeconds += $deltaTimeSeconds
			Write-Log "Failsafe period: $script:g_idleSeconds s < 0" "INFO"
			continue
		}
		
		if ($script:g_idleSeconds -gt $script:Config['IdleSecondsBeforeWeBroadcastSystemIdleEvent'] -and $script:g_isIdle -eq $false) {
			LogSystemEvent_IdleOn
		}
		
		$script:g_nextFileSettingsPoll--
		if ($script:g_nextFileSettingsPoll -le 0 -and -not $script:Config['FileSettingsPollIntervalSeconds'] -eq 0) {
			Update-ConfigFromSettingsFile
			$script:g_nextFileSettingsPoll = $script:Config['FileSettingsPollIntervalSeconds']
		}
		
		# Poll power plan timeout every $script:Config['SettingsPollIntervalSeconds'] updates
		$script:g_nextSettingsPoll--
		if ($script:g_nextSettingsPoll -le 0) {			
			$script:g_isPluggedIn = IsComputerPluggedIn
			if ($script:Config['FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan']) {
				$tuple = Get-PowerPlanIdleTimeoutMinutes $script:g_isPluggedIn
				$newSleepTimeout = $tuple.sleepMinutesVal
				$newHibernateTimeout = $tuple.hibernateMinutesVal
				if ($newSleepTimeout -ne $script:g_CurrentSleepIdleTimeMinutes) {
					Write-Log "Power plan idle timeout changed: $script:g_CurrentSleepIdleTimeMinutes to $newSleepTimeout min." "INFO"
					$script:g_CurrentSleepIdleTimeMinutes = $newSleepTimeout
				}
				if ($newHibernateTimeout -ne $script:g_CurrentSleepIdleTimeMinutes) {
					Write-Log "Power plan idle timeout changed: $script:g_CurrentHibernateIdleTimeMinutes to $newHibernateTimeout min." "INFO"
					$script:g_CurrentHibernateIdleTimeMinutes = $newHibernateTimeout
				}
			}
			
			if ($script:Config['PreventAndReplaceWindowsAutoSleep']) {
				if ($script:Config['TurnOffDisplayAtThisIdleTimeSeconds'] -eq 0) {
					$newDisplayTimeoutDurationSeconds = Get-DisplayTimeoutSeconds $script:g_isPluggedIn
					if($newDisplayTimeoutDurationSeconds -ne $script:g_DisplayTimeoutDurationSeconds) {
						Write-Log "  System g_DisplayTimeoutDurationSeconds changed: $script:g_DisplayTimeoutDurationSeconds to $newDisplayTimeoutDurationSeconds seconds." "INFO"
						$script:g_DisplayTimeoutDurationSeconds = $newDisplayTimeoutDurationSeconds
					}
				}
				
				if ($script:Config['TurnOnScreensaverAtThisIdleTimeSeconds'] -eq 0) {
					$newScreensaverTimeoutDurationSeconds = Get-ScreensaverTimeoutSeconds
					if($newScreensaverTimeoutDurationSeconds -ne $script:g_ScreensaverTimeoutDurationSeconds) {
						Write-Log "  System g_ScreensaverTimeoutDurationSeconds changed: $script:g_ScreensaverTimeoutDurationSeconds to $newScreensaverTimeoutDurationSeconds seconds." "INFO"
						$script:g_ScreensaverTimeoutDurationSeconds = $newScreensaverTimeoutDurationSeconds
					}
				}
			}
			
			$script:g_nextSettingsPoll = $script:Config['SettingsPollIntervalSeconds']
		}
		
		if ((Test-Path $script:Config['PauseFlagPath']) -eq $true -or $script:Config['PauseScript'] -eq $true) {
			$script:g_idleSeconds = 0.0
			if ($script:g_isIdle -eq $true) {
				LogSystemEvent_IdleOff
			}
			if ($script:g_PreventSleep_ES -eq $true -and $script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true){
				$script:g_PreventSleep_ES = $false
				[WindowsSleepWrangler]::StopIgnoringIdleTimers()
			}
			continue
		} elseif ($script:g_PreventSleep_ES -eq $false -and $script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true) {
			$script:g_PreventSleep_ES = $true
			[WindowsSleepWrangler]::IgnoreIdleTimers()
		}
		
		# CPU
		$cpu = Get-CpuUsagePercent
		$cpuAbove = $cpu -gt $script:Config['CpuThresholdPercent']
		$script:g_cpuHistory.Enqueue($cpuAbove)
		if ($script:g_cpuHistory.Count -gt $script:g_maxSamples) { 
			$null = $script:g_cpuHistory.Dequeue()
		}

		# Disk
		$disk = Get-DiskIoKBps
		$diskAbove = $disk -gt $script:Config['DiskThresholdKBps']
		$script:g_diskHistory.Enqueue($diskAbove)
		if ($script:g_diskHistory.Count -gt $script:g_maxSamples) { 
			$null = $script:g_diskHistory.Dequeue()
		}

		# Network, sums up download and upload on all active non-virtual network interfaces
		$net = Get-NetworkIoKBps
		#Write-Log "net: $net" "DEBUG"
		$netAbove = $net -gt $script:Config['NetworkThresholdKBps']
		$script:g_netHistory.Enqueue($netAbove)
		if ($script:g_netHistory.Count -gt $script:g_maxSamples) { 
			$null = $script:g_netHistory.Dequeue()
		}
		
		# Audio
		$audioPlaying = Get-AudioIsPlaying
		$script:g_audioHistory.Enqueue($audioPlaying)
		if ($script:g_audioHistory.Count -gt $script:g_maxSamplesAudio) { 
			$null = $script:g_audioHistory.Dequeue()
		}

		# Mouse, replaceed with Get-SecondsSinceLastInputInfo
		#$currMouse = Get-MouseMovementPixels
		#$mouseDelta = [math]::Sqrt((($currMouse.X - $prevMouse.X) * ($currMouse.X - $prevMouse.X)) + 
		#							(($currMouse.Y - $prevMouse.Y) * ($currMouse.Y - $prevMouse.Y)))
		#$prevMouse = $currMouse
		#$mouseMoved = $mouseDelta -gt $MouseThresholdPixels
		
		# Check for sustained activity (reset idle timer)
		$hasSustainedActivity = $false
		
		# CPU: >= $script:Config['ActiveSamplesWithinInterval'] of last $script:g_maxSamples were above threshold
		if ($script:g_cpuHistory.Count -ge $script:g_maxSamples) {
			$activeCount = ($script:g_cpuHistory | Where-Object { $_ }).Count
			if ($activeCount -ge $script:Config['ActiveSamplesWithinInterval']) {
				if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
					Write-Log "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] CPU: $activeCount/$script:g_maxSamples samples > $($script:Config['CpuThresholdPercent'])% (>= $($script:Config['ActiveSamplesWithinInterval']) required) for $($script:Config['ActivityDetectionPeriodSeconds']) sec" "INFO"
				} elseif ($script:Config['LogToConsoleVerbose']) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] CPU: $activeCount/$script:g_maxSamples samples > $($script:Config['CpuThresholdPercent'])% (>= $($script:Config['ActiveSamplesWithinInterval']) required) for $($script:Config['ActivityDetectionPeriodSeconds']) sec" "INFO"
				}
				$hasSustainedActivity = $true
			}
		}

		# Disk: >= $script:Config['ActiveSamplesWithinInterval'] of last $script:g_maxSamples were above threshold
		if ($script:g_diskHistory.Count -ge $script:g_maxSamples) {
			$activeCount = ($script:g_diskHistory | Where-Object { $_ }).Count
			if ($activeCount -ge $script:Config['ActiveSamplesWithinInterval']) {
				if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
					Write-Log "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Disk: $activeCount/$script:g_maxSamples samples > $($script:Config['DiskThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required) for $($script:Config['ActivityDetectionPeriodSeconds']) sec" "INFO"
				} elseif ($script:Config['LogToConsoleVerbose']) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Disk: $activeCount/$script:g_maxSamples samples > $($script:Config['DiskThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required) for $($script:Config['ActivityDetectionPeriodSeconds']) sec" "INFO"
				}
				$hasSustainedActivity = $true
			}
			
		}

		# Network: >= $script:Config['ActiveSamplesWithinInterval'] of last $script:g_maxSamples were above threshold
		if ($script:g_netHistory.Count -ge $script:g_maxSamples) {
			$activeCount = ($script:g_netHistory | Where-Object { $_ }).Count
			if ($activeCount -ge $script:Config['ActiveSamplesWithinInterval']) {
				if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
					Write-Log "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Network: $activeCount/$script:g_maxSamples samples > $($script:Config['NetworkThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required) for $($script:Config['ActivityDetectionPeriodSeconds']) sec" "INFO"
				} elseif ($script:Config['LogToConsoleVerbose']) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Network: $activeCount/$script:g_maxSamples samples > $($script:Config['NetworkThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required) for $($script:Config['ActivityDetectionPeriodSeconds']) sec" "INFO"
				}
				$hasSustainedActivity = $true
			}
		}

		# Audio: all samples must be true
		if ($script:g_audioHistory.Count -eq $script:g_maxSamplesAudio -and ($script:g_audioHistory | Where-Object { $_ }).Count -eq $script:g_maxSamplesAudio) {
			if($script:g_idleSeconds -ge $script:Config['ActivityDetectionPeriodSecondsAudio']){
				if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
					Write-Log "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Sustained audio playing for $($script:Config['ActivityDetectionPeriodSecondsAudio']) sec, Resetting idle counter." "INFO"
				} elseif ($script:Config['LogToConsoleVerbose']) {
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Sustained audio playing for $($script:Config['ActivityDetectionPeriodSecondsAudio']) sec, Resetting idle counter." "INFO"
				}
				$hasSustainedActivity = $true
			}
		}
		
		#if ($mouseMoved) {
		$inputBasedActivityThisFrame = $false
		$secondsSinceLastInputInfo = Get-SecondsSinceLastInputInfo
		if ($secondsSinceLastInputInfo -and $secondsSinceLastInputInfo -le $script:g_idleSeconds) {
			#if ($script:g_idleSeconds -ge $script:Config['ActivityDetectionPeriodSeconds']) {
				if($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']){
					#Write-Log "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Mouse moved $([math]::Round($mouseDelta,1)) px > $MouseThresholdPixels, Resetting idle counter." "INFO"
					Write-Log "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Mouse/touch/keyboard activity registered $secondsSinceLastInputInfo seconds ago. Resetting idle counter." "INFO"
				} elseif ($script:Config['LogToConsoleVerbose']) {
					#Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Mouse moved $([math]::Round($mouseDelta,1)) px > $MouseThresholdPixels, Resetting idle counter." "INFO"
					Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $script:g_idleSeconds] Mouse/touch/keyboard activity registered $secondsSinceLastInputInfo seconds ago. Resetting idle counter." "INFO"
				}
				$hasSustainedActivity = $true
				$inputBasedActivityThisFrame = $true
			#}
			
			if ($script:Config['PreventAndReplaceWindowsAutoSleep']) {
				$script:g_ScreenSaverStarted = $false
				$script:g_DisplayTurnedOff = $false
			} elseif ($FollowTheSameSleepTimeSettingAsYourPowerPlan -eq $false) {
				if($script:Config['TurnOnScreensaverAtThisIdleTimeSeconds'] -gt 0) {
					$script:g_ScreenSaverStarted = $false
				}
				if($script:Config['TurnOffDisplayAtThisIdleTimeSeconds'] -gt 0) {
					$script:g_DisplayTurnedOff = $false
				}
			}
			
			if($script:Config['LockPcAtThisIdleTimeSeconds'] -gt 0) {
				$script:g_PcLockedOnDemand = $false
			}
		}
		
		
		# Log to file sometimes, and to console some other times
		if (($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds'] -and $hasSustainedActivity -eq $true) -or ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds'] -and $script:g_nextSettingsPoll -le $script:Config['ActivityDetectionPeriodSeconds'])) {
			$mouseLog = " "
			if (Test-IsInteractiveSession) {
				#$mouseLog = "MouseDelta: $([math]::Round($mouseDelta,1)) px"
				$mouseLog = "Input: $secondsSinceLastInputInfo s ago"
			} else {
				$mouseLog = "(mouse check skipped)"
			}
			$statusMessage = "[idleSeconds: $script:g_idleSeconds] CPU: $cpu % | Disk: $disk KBps | Net: $net KBps | $mouseLog | Idle: $([math]::Round($script:g_idleSeconds/60,2)) min | (T Sleep: $script:g_CurrentSleepIdleTimeMinutes min | T Hibernate: $script:g_CurrentHibernateIdleTimeMinutes min | T Display: $script:g_DisplayTimeoutDurationSeconds sec | T ScreenSaver: $script:g_ScreensaverTimeoutDurationSeconds sec)"
			Write-Log $statusMessage "INFO"
		} elseif ($script:Config['LogToConsoleVerbose']) {
			$mouseLog = " "
			if (Test-IsInteractiveSession) {
				#$mouseLog = "MouseDelta: $([math]::Round($mouseDelta,1)) px"
				$mouseLog = "Input: $secondsSinceLastInputInfo s ago"
			} else {
				$mouseLog = "(mouse check skipped)"
			}
			$statusMessage = "[idleSeconds: $script:g_idleSeconds] CPU: $cpu % | Disk: $disk KBps | Net: $net KBps | $mouseLog | Idle: $([math]::Round($script:g_idleSeconds/60,2)) min | (T Sleep: $script:g_CurrentSleepIdleTimeMinutes min | T Hibernate: $script:g_CurrentHibernateIdleTimeMinutes min | T Display: $script:g_DisplayTimeoutDurationSeconds sec | T ScreenSaver: $script:g_ScreensaverTimeoutDurationSeconds sec)"
			Write-Host-Wrapper $statusMessage "INFO"
		}

		if ($hasSustainedActivity) {
			$script:g_idleSeconds = 0.0
			if ($script:g_isIdle -eq $true) {
				LogSystemEvent_IdleOff
			}
			$script:g_cpuHistory.Clear()
			$script:g_diskHistory.Clear()
			$script:g_netHistory.Clear()
			$script:g_audioHistory.Clear()
		} else {
			$script:g_idleSeconds += $deltaTimeSeconds
		}

		if($inputBasedActivityThisFrame -eq $false) {
			# Check if we are in charge of turning off the display or turning on any screensaver, and do it if it's time
			if ($script:Config['PreventAndReplaceWindowsAutoSleep']) {
				if($script:g_ScreenSaverStarted -eq $false -and $script:g_ScreensaverTimeoutDurationSeconds -and $script:g_ScreensaverTimeoutDurationSeconds -gt $script:Config['FailsafeTimeSeconds'] -and $script:g_idleSeconds -gt $script:g_ScreensaverTimeoutDurationSeconds) {
					Start-Screensaver
				}
				
				if($script:g_DisplayTurnedOff -eq $false -and $script:g_DisplayTimeoutDurationSeconds -and $script:g_DisplayTimeoutDurationSeconds -gt $script:Config['FailsafeTimeSeconds'] -and $script:g_idleSeconds -gt $script:g_DisplayTimeoutDurationSeconds) {
					Turn-Display-Off
				}
			} elseif ($FollowTheSameSleepTimeSettingAsYourPowerPlan -eq $false) {
				if($script:Config['TurnOnScreensaverAtThisIdleTimeSeconds'] -gt 0 -and $script:g_ScreenSaverStarted -eq $false -and $script:g_ScreensaverTimeoutDurationSeconds -and $script:g_ScreensaverTimeoutDurationSeconds -gt $script:Config['FailsafeTimeSeconds'] -and $script:g_idleSeconds -gt $script:g_ScreensaverTimeoutDurationSeconds) {
					Start-Screensaver
				}
				
				if($script:Config['TurnOffDisplayAtThisIdleTimeSeconds'] -gt 0 -and $script:g_DisplayTurnedOff -eq $false -and $script:g_DisplayTimeoutDurationSeconds -and $script:g_DisplayTimeoutDurationSeconds -gt $script:Config['FailsafeTimeSeconds'] -and $script:g_idleSeconds -gt $script:g_DisplayTimeoutDurationSeconds) {
					Turn-Display-Off
				}
			}
			# else NOTE: if $FollowTheSameSleepTimeSettingAsYourPowerPlan is true and $script:Config['PreventAndReplaceWindowsAutoSleep'] is false, then the PC will turn off its display and turn on its screensaver on its own, no involvement from us
		}
		
		if ($script:g_PcLockedOnDemand -eq $false -and $script:Config['LockPcAtThisIdleTimeSeconds'] -and $script:Config['LockPcAtThisIdleTimeSeconds'] -gt $script:Config['FailsafeTimeSeconds'] -and $script:g_idleSeconds -gt $script:Config['LockPcAtThisIdleTimeSeconds']) {
			# Write-Log "g_PcLockedOnDemand: $script:g_PcLockedOnDemand, LockPcAtThisIdleTimeSeconds: $script:Config['LockPcAtThisIdleTimeSeconds'], failsafeTimeSeconds: $script:Config['FailsafeTimeSeconds'], idleSeconds: $script:g_idleSeconds" "INFO"
			Lock-PC
		}

		# Note: this doesn't work unless you run the script as administrator, so I commented it out ctrl+f:[respectOtherApps]
		#if ($RespectOtherAppsSleepExecutionPreventionFlags -eq $false -or (Test-OtherSystemExecutionStateHeld -eq $false -and $RespectOtherAppsSleepExecutionPreventionFlags -eq $true)) {
			# Check if ready to sleep
			if ($script:g_idleSeconds -ge ($script:g_CurrentSleepIdleTimeMinutes * 60)) {
				$abort = Show-AbortDialog -Seconds $AbortWindowCountdownSeconds

				if ($abort) {
					Write-Log "User aborted sleep." "INFO"
					$script:g_idleSeconds = 0.0
					if ($script:g_isIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
				} else {
					Write-Log "Proceeding to sleep because no abort or non-interactive session..." "INFO"
					#if (Test-Path $script:Config['PauseFlagPath']) {
					#    Remove-Item $script:Config['PauseFlagPath'] -Force
					#    Write-Log "Deleted flag file: $script:Config['PauseFlagPath']" "INFO"
					#}
					
					Enter-SleepState
					$script:g_idleSeconds = 0.0
					if ($script:g_isIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
					Write-Log "System be woke. Resuming monitoring." "INFO"
				}
			}
			# Check if ready to hibernate
			if ($script:g_idleSeconds -ge ($script:g_CurrentHibernateIdleTimeMinutes * 60)) {
				$abort = Show-AbortDialog -Seconds $AbortWindowCountdownSeconds

				if ($abort) {
					Write-Log "User aborted hibernate." "INFO"
					$script:g_idleSeconds = 0.0
					if ($script:g_isIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
				} else {
					Write-Log "Proceeding to hibernate because no abort or non-interactive session..." "INFO"
					#if (Test-Path $script:Config['PauseFlagPath']) {
					#    Remove-Item $script:Config['PauseFlagPath'] -Force
					#    Write-Log "Deleted flag file: $script:Config['PauseFlagPath']" "INFO"
					#}
					
					Enter-HibernateState
					$script:g_idleSeconds = 0.0
					if ($script:g_isIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
					Write-Log "System be woke. Resuming monitoring." "INFO"
				}
			}
		#}
	}
} finally {
	if($script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true){
		$script:g_PreventSleep_ES = $false
		[WindowsSleepWrangler]::StopIgnoringIdleTimers()
	}
	if ($script:g_isIdle -eq $true) {
		LogSystemEvent_IdleOff
	}
}