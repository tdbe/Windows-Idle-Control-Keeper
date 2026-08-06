<#
	MIT License, Copyright (c) 2026 Tudor Berechet [tdbe](https://github.com/tdbe) 
	
.SYNOPSIS

	🕯️W.I.C.K. Handles system Idle using your specific thresholds & conditions. It triggers or prevents Windows sleep, hibernate, display, lock, screen saver, on Your terms. Detects activity on CPU, GPU, (non-virtual) Network (internet & LAN), Storage, Input, peak Sound value. Idle-breaking events via frequency & amplitude / time periods. Power Plan aware.
	
.DESCRIPTION

	# Windows-Idle-Control-Keeper

	# Intro:

	[Saul Goodman voice] **Can't get reliable sleep? Feeling like it's out of your control? Well fret not, just run this script and you can start counting those sheep!**

	**WICK - Windows Idle Control Keeper**

	This non-admin script detects Idle activity with your specific thresholds and conditions, and triggers / prevents Windows sleep, hibernate, display, lock, screen saver, on Your terms. Detects activity on CPU, GPU, (non-virtual) Network (both internet and LAN), Storage, input, and peak Sound value; counting idle-breaking event frequency and amplitude per time periods, to determine if an Idle timer should continue or be broken. It's Windows Power Plan aware, including display off and screensaver schedule, and maintains windows screen locking.

	I don't usually post my system scripts but it annoyed me that for such a wide need, there was nothing out there but forum threads of people using obscure and partial tools like [DontSleep!.exe](https://www.softwareok.com/?Download=DontSleep) [from 2014](https://www.chip.de/downloads/Don-t-Sleep_42626965.html)

	# Features, Dependencies, Log Example, Notes, Run & Parameters:

	## Features:

	- Does not require administrator permission.
	- Works even if windows is locked. 
	- Also works as a system task: if you start it at system start via task scheduler + `run whether the user is logged in or not`. But then you must also add a user task on `log in` so it starts to monitor your actual power plan: read the description of the `MutualExclusionFlagFile` parameter. (And if for some reason you Log Off but keep the computer on, you need to manually start a new system task.)
	- Shows warning / abort window for AbortWindowCountdownSeconds (and CriticalBatteryAbortWindowCountdownSeconds) before triggering a sleep or hibernate (if in an interactive session (not locked or logged out)).
	- Also provides its own options to sleep or hibernate when below specific battery levels (if there is a battery): SleepWhenBatteryIsBelowLevel, HibernateWhenBatteryIsBelowLevel. A self-dismissing message is shown to abort if session is interactive (CriticalBatteryAbortWindowCountdownSeconds).
	- Dynamically reads (every minute (configurable)) from your currently active windows power plan (plugged in or battery) to check sleep and hibernate times (also display and screensaver) (can also ignore them and use manual times).
	- Also can read from Settings_File_Windows_Idle_Control_Keeper_txt. So you can pause/resume, or tweak settings, while the script is running (every FileSettingsPollIntervalMinutes) (there's a flag: -IgnoreTheSettingsFile).
	- Determines idle by accumulating sustained activity event samples, of configurable frequency & amplitude, over certain timeframes (and uses delta time), based on: if there's CPU, GPU, Network, sorage (without waking sleeping hard disks), audio spikes, and input activity.
	- Can prevent windows from sleeping/hibernate until this script decides it's time, or work alongside it.
	- Can set a sleep or hibernate time for longer than 5h (the max that Windows power plan allows for some gormless reason).
	- Allows a blacklist for logical drives e.g. `"L", "A", "N"` - you may have drives that have activity you consider passive and you're okay sleeping on. But also keep in mind the NetworkThresholdKBps setting.
	- Can also prevent sleep or hibernate while a file or folder exists (e.g. "D:\Bak" or "D:\Bak\myLongSlowBackup.tmp"), by pasting its path in the DontSleepWhileThisFileExistsPath. If this file exists in the moment that this script would trigger a sleep or hibernate, it won't sleep/hibernate. (File checked every FileSettingsPollIntervalMinutes) (The script otherwise works as normal e.g. turn display off (if set to)) The path should be accessible by all users.
	- Logs what's going on, to Windows' Event Viewer - Application Log (0. on start ('69' '1111'), 1. idle_on ('69' '420') (after 1m of idle), 2. idle_off ('69' '421'), 3. on exit ('69' '1000')) (the actual message is in the 'Details' tab of the event (non-admin limitation)). Also logs to file at LogPath, so you know at what time of day Idle state was broken, by what, and after how much idle time. (or if there were errors) (log cleans itself up to stay less than LogMaxSizeMB)
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
	[2026-04-30 00:11:58] [INFO] Log path: C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log
	[2026-04-30 00:11:58] [INFO] Settings path: C:\Commands_And_Logs\[Settings_File]_Windows_Idle_Control_Keeper.txt
	[2026-04-30 00:11:58] [INFO] You set to use windows power plan's sleep and hiberante values (if sleep/hibernate are enabled on the system): 30 min and 60 min. (We check to update this value every: SettingsPollIntervalMinutes: 1 min.)
	[2026-04-30 00:11:58] [INFO] It's been 171368.983134705 minute(s) since the last update, which means we
	were sleeping or somehow lagging a lot, Resetting idle counter.
	[2026-04-30 00:13:38] [INFO][IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required). [idleSeconds: 1.97975]][deltaTime: 1.81188]
	[2026-04-30 00:15:49] [INFO][IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required). [idleSeconds: 131.58716][deltaTime: 1.79505]
	[2026-04-30 00:18:03] [INFO][IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required). [idleSeconds: 134.36724][deltaTime: 1.80448]
	[2026-04-30 00:20:18] [INFO][IDLE BREAKER] Sustained audio playing for 5 samples. Resetting idle counter. [idleSeconds: 135.39535][deltaTime: 1.83311]
	[2026-04-30 00:22:27] [INFO][IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required). [idleSeconds: 129.7384][deltaTime: 1.80238]
	[2026-04-30 00:24:35] [INFO][IDLE BREAKER] Network: 3/6 samples > 850 KBps (>= 3 required). [idleSeconds: 128.48579][deltaTime: 1.81196]
	[2026-04-30 00:26:53] [INFO][IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required). [idleSeconds: 138.36093][deltaTime: 1.80977]
	[2026-04-30 00:29:07] [INFO][IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required). [idleSeconds: 134.53758][deltaTime: 1.80943]
	[2026-04-30 00:31:16] [INFO][IDLE BREAKER] Sustained audio playing for 5 samples. Resetting idle counter. [idleSeconds: 129.49264][deltaTime: 1.82225]
	[2026-04-30 00:33:22] [INFO][IDLE BREAKER] Disk: 4/6 samples > 1250 KBps (>= 3 required). [idleSeconds: 126.15309][deltaTime: 1.80812]
	[2026-04-30 00:35:41] [INFO][IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required). [idleSeconds: 138.73624][deltaTime: 1.81147]
	[2026-04-30 00:38:00] [INFO][IDLE BREAKER] Network: 3/6 samples > 850 KBps (>= 3 required). [idleSeconds: 139.63062][deltaTime: 1.80641]
	[2026-04-30 00:42:30] [INFO][IDLE BREAKER] Network: 5/6 samples > 850 KBps (>= 3 required). [idleSeconds: 270.65825][deltaTime: 1.81357]
	[2026-04-30 00:44:38] [INFO][IDLE BREAKER] Mouse/touch/keyboard activity registered 1 seconds ago. Resetting idle counter. [idleSeconds: 128.70253][deltaTime: 1.81204]
	[2026-04-30 00:46:52] [INFO][IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required). [idleSeconds: 134.8361][deltaTime: 1.80762]
	[2026-04-30 00:49:18] [INFO][IDLE BREAKER] Network: 4/6 samples > 850 KBps (>= 3 required). [idleSeconds: 146.67395][deltaTime: 1.81792]
	[2026-04-30 00:51:39] [INFO][IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required). [idleSeconds: 141.39137][deltaTime: 1.81315]
	[2026-04-30 00:53:57] [INFO][IDLE BREAKER] Network: 6/6 samples > 850 KBps (>= 3 required). [idleSeconds: 138.73501][deltaTime: 1.82582]
	[2026-04-30 00:53:59] [INFO] CPU: 1 % | GPU: 6 % | Disk: 8279 KBps | Net: 1201 KBps | Input: 91 s ago | Idle: 0.0 min | (T Sleep: 30 min | T Hibernate: 60 min | T Display: 20.0 min | T ScreenSaver: 0.0 min | T Demand Win Lock: 15.0 min). [idleSeconds: 2.52849][deltaTime: 1.83315]
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

	I wouldn't be caught dead writing for free 2000 lines of powershell script of all things, so for this I tried out LLMs. I don't consider llm output even remotely reliable, but this is all verified and very re-written by me. For those curious: I used 256k context and: qwen 3 coder next 80b a3b q6, qwen 3.6 35b a3b q8, and qwen 3.6 27b q4, locally. They're "great" (within 5-10% of the scores of the huge frontier models) but simultaneously also completely shit at even such a simple job, and not just because this solution doesn't already exist: ie they picked network and storage checks that take at least 1s to return a value, and were calling them repeatedly in loops per disk and per adapter, resulting in a while loop that runs once every 7-10s.. So the verdict is I had to do all the thinking myself. They only oneshotted the logging, the cpu, the sleep functions, and mostly the .PARAM & settings handling. Also the peak audio checking I had to research and write myself in python after many wildly off LLM solutions.

	## Run & Parameters:

	There are a LOT of parameters you can set when calling the script or adding it to Task Scheduler (the ones you skip will have defaults).

	To see all parameters and their description, run this command: 

	```
	get-help "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -detailed
	```

	### Run in a powershell terminal window, examples:

	```
	powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$true -PreventAndReplaceWindowsAutoSleep:$true -IgnoreTheSettingsFile:$true # other flags -etc. -etc.

	or

	powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$false -UserSpecifiedSleepIdleTimeMinutes:720 -PreventAndReplaceWindowsAutoSleep:$true -LockPcAtThisIdleTimeMinutes:3 -TurnOnScreensaverAtThisIdleTimeMinutes:0 -TurnOffDisplayAtThisIdleTimeMinutes:6 -IgnoreTheSettingsFile:$true # other flags -etc. -etc.

	```

	### Run in Task Scheduler:

	#### Program/script: 

	```
	powershell.exe
	```

	#### Add arguments (window opens as minimized):

	```
	-NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$true -UserSpecifiedSleepIdleTimeMinutes:30 -PreventAndReplaceWindowsAutoSleep:$true -IgnoreTheSettingsFile:$true # other flags -etc. -etc.
	```

	#### Add arguments (no window, runs in background completely hidden):

	```
	-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$true -UserSpecifiedSleepIdleTimeMinutes:30 -PreventAndReplaceWindowsAutoSleep:$true -IgnoreTheSettingsFile:$true # other flags -etc. -etc.
	```
	
	### Alternative in case window is not hidden
	
	#### Program/script: 

	```
	C:\Windows\System32\conhost.exe
	```
	
	#### Add arguments (no window, runs in background completely hidden):

	```
	--headless powershell -WindowStyle hidden -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan:$true -UserSpecifiedSleepIdleTimeMinutes:30 -PreventAndReplaceWindowsAutoSleep:$true -IgnoreTheSettingsFile:$true # other flags -etc. -etc.
	```
	
	### PS:
	
	Instead of regular params, I added $script:Config that loads from either the params ($PSBoundParameters) or the Settings_File_Windows_Idle_Control_Keeper_txt. So you can pause or tweak settings while the script is running (every FileSettingsPollIntervalMinutes).
	
.PARAMETER IgnoreTheSettingsFile
  The entries in the settings file, even if default values, will overwrite the command line parameters unless the command line has the -IgnoreTheSettingsFile:$true (it will always still read the PauseScript parameter). The settings file is read every FileSettingsPollIntervalMinutes. (default: $false)
  
.PARAMETER PauseScript
  You can pause and resume the script in real time (every FileSettingsPollIntervalMinutes) using this parameter in the settings file. From the cli it will simply start the script as paused (the settings file is still polled while paused). NOTE: This param gets read even if -IgnoreTheSettingsFile:$true. (default: $false)

.PARAMETER PreventAndReplaceWindowsAutoSleep
  Uses `SetThreadExecutionState` to prevent idle-based sleep commands. Note that actively triggering Sleep e.g. via the Start menu, or via an explicit function call (e.g. SetSuspendState) from some active software, or laptop lid close, will STILL cause the PC to sleep! Also, if this is set to $false, you will get windows' event, plus also this script's event (e.g. windows turns display off (if set to), and also this script turns display off (if set to) - so, whichever comes first). (default: $true)
  
.PARAMETER FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan
  Read idle timeout values from Windows power plan (sleep, hibernate, display). Can still be overridden by the user defined sleep, hibernate, display off, screensaver time, if set to greater than zero (and whichever comes first will be triggered, unless you PreventAndReplaceWindowsAutoSleep:$true). (default: $true)

.PARAMETER UserSpecifiedSleepIdleTimeMinutes
  Using this you can set more than the weird 5h max (e.g. 720 mins (12 hours)) limit that Windows power plan lets you set.  
  If nonzero, overrides itself even if the FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan == $true. (default: 0)
  
.PARAMETER UserSpecifiedHibernateIdleTimeMinutes
  Using this you can set more than the weird 5h max (e.g. 720 mins (12 hours)) limit that Windows power plan lets you set.  
  If nonzero, overrides itself even if the FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan == $true. (default: 0)
  
.PARAMETER TurnOnScreensaverAtThisIdleTimeMinutes
  Zero means it doesn't trigger (or the system screensaver setting is used). You can also use decimals e.g. '0.3' min. (default: 0) 
  
.PARAMETER TurnOffDisplayAtThisIdleTimeMinutes
  If nonzero, overrides itself even if the FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan == $true. Zero means it doesn't trigger (or the system power setting is used). You can also use decimals e.g. '0.3' min. (default: 0) 

.PARAMETER LockPcAtThisIdleTimeMinutes 
  If not zero ((and after the FailsafeTimeMinutes (which is active on startup or on resume from sleep)) it will lock pc at this idle time, which you can set to be earlier than whenever Windows decides to lock. You can also use decimals e.g. '0.3' min. (default: 0)

.PARAMETER CpuThresholdPercent
  CPU usage above this resets idle timer. (default: 7)
  
.PARAMETER GpuThresholdPercent
  GPU usage above this resets idle timer. (default: 30)

.PARAMETER DiskThresholdKBps
  Disk I/O (KB/s) above this resets idle timer. Sums (read + write) * all the disks except the DiskBlacklistDrives. It also detects any ongoing volume shadowcopy service activity e.g. if you're doing a backup. (default: 1250)

.PARAMETER NetworkThresholdKBps
  Network I/O (KB/s) above this resets idle timer. Sums (download + upload) * all the non-virtual network interfaces. (default: 850)

.PARAMETER ActiveSamplesWithinInterval
  How many instances (seconds, but there can be lag) of activity must be detected within the last ActivityDetectionPeriodSamples seconds for us to consider that activity an idle breaker. (default: 3)

.PARAMETER ActivityDetectionPeriodSamples
  Seconds window to check for sustained activity. (default: 6)

.PARAMETER ActivityDetectionPeriodSamplesAudio
  Separate timeout for audio - counts if there was constant sound for this many samples in a row; with a custom sound peak threshold to e.g. avoid background noise (threshold in audio pythin script). (default: 5)
  
.PARAMETER UseOnlyInputAndAudioEventsForDisplayOff
  If set to $true, the screensaver (if available) and the monitor off, will happen based on a separate idle timer which can only be interrupted by audio playback or user input. This is nice if you want to put your computer to work, and walk away with the guarantee that the monitor won't waste power even if the pc will not sleep. (default: $false)

.PARAMETER IdleSecondsBeforeWeBroadcastSystemIdleEvent
  How much idle time must pass before we declare the system idle as far as this script is concerned, and send an idle event to the windows Event Viewer's Application Log (regardless of when the display is turned off or screensaver turns on or anything else). (default: 60)

.PARAMETER DiskBlacklistDrives
  Allows blacklist for logical drives e.g. `"L", "A", "N"` - drives that have activity but you consider passive and you're okay sleeping on them. Keep in mind the NetworkThresholdKBps setting.  
  (default: @("E", "F"))
  
.PARAMETER SleepWhenBatteryIsBelowLevel
  If you don't have a battery, it will count the battery level as always at 100. So 0 means disabled (can't be less than 0). (default: 0)

.PARAMETER HibernateWhenBatteryIsBelowLevel
  If you don't have a battery, it will count the battery level as always at 100. So 0 means disabled (can't be less than 0). (default: 0)
  
.PARAMETER Settings_File_Windows_Idle_Control_Keeper_txt
  These settings can be edited while the script is running and the script will read them every FileSettingsPollIntervalMinutes. The path should be accessible by all users. (default: "C:\Commands_And_Logs\[Settings_File]_Windows_Idle_Control_Keeper.txt")
  
.PARAMETER PycawAudioCheckerPath
  Full path to the Python script used to detect audio playback with custom threshold. The path should be accessible by all users.
  (default: "C:\Commands_And_Logs\Pycaw_check_if_audio_is_playing.py")

.PARAMETER PythonPath
  Full path to the Python executable used to run the audio checker script. It detects the latest version on its own if you use that wildcard.
  (default: "$env:USERPROFILE\AppData\Local\Programs\Python\Python{VERSION}\python.exe")

.PARAMETER DontSleepWhileThisFileExistsPath
  You can prevent sleep or hibernate while a file (or directory) path exists (e.g. "D:\Bak" or "D:\Bak\myLongSlowBackup.tmp"), by pasting its path in the DontSleepWhileThisFileExistsPath. If this file exists in the moment that this script would trigger a sleep or hibernate, it won't sleep/hibernate. (File checked every FileSettingsPollIntervalMinutes) (The script otherwise works as normal e.g. turn display off (if set to)) The path should be accessible by all users.
  (default: "D:\Bak\myLongSlowBackup.tmp")

.PARAMETER MutualExclusionFlagFile
  This flag file is created by the script on start. The path should be accessible by all users. It's used to make sure that if you start a new instance of this script, the old instance sees the new PID in this flag file, and exits. This is important also if you need to run the script before any user is logged in: if you "run whether the user is logged in or not" then the task scheduler makes the script always run under the SYSTEM account - which means it will only get the hidden system user's power plan settings for display and sleep etc. What you can do is start 2 tasks: 1 as system at startup, and 1 on user log in (log in, not unlock). The newer log-in task script will cause the old script(s) to exit.
  (default: "C:\Commands_And_Logs\.WickMutualExclusionFlag")

.PARAMETER LogPath
  Full path to the log file. The path should be accessible by all users. (default: "C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log")

.PARAMETER LogMaxAgeDays
  Keep logs this many days. (default: 30)

.PARAMETER LogMaxSizeMB
  Rotate log if larger than this MB. (default: 10)

.PARAMETER LogToFileIntervalSeconds
  To prevent writing to file every second while you're using the PC, it won't log unless it's been idle for this many seconds. (default: 60)
  
.PARAMETER LogToConsoleVerbose
  Whether to log to the console (not log file) as often as there is an event in the constant loop. (default: true)

.PARAMETER AbortWindowCountdownSeconds
  Seconds to show the sleep or hibernate abort dialog before triggering sleep or hibernate. (default: 60)
  
.PARAMETER CriticalBatteryAbortWindowCountdownSeconds
  This should be short so the computer doesn't drain while waiting on the message. Seconds to show the sleep or hibernate abort dialog before triggering sleep or hibernate. (default: 60)
  
.PARAMETER SettingsPollIntervalMinutes
  Value should be lower than your e.g. sleep time (and in some sort of tandem with `FileSettingsPollIntervalMinutes`). Dynamically updates various script settings and timers, including checking the MutualExclusionFlagFile, or settings from your currently active windows power plan (plugged in or battery) to check sleep and also hibernate times. You can also use decimals e.g. '0.3' min. (default: 1)
  
.PARAMETER FileSettingsPollIntervalMinutes
  Dynamically reads parameters from your settings file (should run at same interval as SettingsPollIntervalMinutes, unless you're not using it (0 means not used)). You can also use decimals e.g. '0.3' min. (default: 1)
  
.PARAMETER FailsafeTimeMinutes
  When the script starts or when it resumes from sleep, I use this failsafe timer, in case somebody screws something up / adds something that for example would lock the pc every second. This way if you sleep + wake, or restart the pc, you get e.g. 60 seconds to stop it even if you set it to run hidden on system startup from task scheduler. (default: 0.98)
#>

# Note: this doesn't work unless you run the script as administrator, so I commented it out ctrl+f:[respectOtherApps]
#.PARAMETER RespectOtherAppsSleepExecutionPreventionFlags
#  Respects other apps if/when they do what we ourselves do with the -PreventAndReplaceWindowsAutoSleep flag. (default: $false)

# ───────────────────────────────────────────────────────────────────────────────
# 1. PARAMETER BLOCK (CLI overrides only)
# ───────────────────────────────────────────────────────────────────────────────
param(
	[bool]$PauseScript,
	[bool]$IgnoreTheSettingsFile,
    [bool]$PreventAndReplaceWindowsAutoSleep,
    [bool]$FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan,
    [int]$UserSpecifiedSleepIdleTimeMinutes,
    [int]$UserSpecifiedHibernateIdleTimeMinutes,
    [int]$TurnOnScreensaverAtThisIdleTimeMinutes,
    [int]$TurnOffDisplayAtThisIdleTimeMinutes,
    [int]$CpuThresholdPercent,
    [int]$GpuThresholdPercent,
    [int]$DiskThresholdKBps,
    [int]$NetworkThresholdKBps,
    [int]$ActiveSamplesWithinInterval,
    [int]$ActivityDetectionPeriodSamples,
    [int]$ActivityDetectionPeriodSamplesAudio,
    [int]$LockPcAtThisIdleTimeMinutes,
    [int]$IdleSecondsBeforeWeBroadcastSystemIdleEvent,
    [bool]$UseOnlyInputAndAudioEventsForDisplayOff,
    [string[]]$DiskBlacklistDrives,
	[int]$SleepWhenBatteryIsBelowLevel,
	[int]$HibernateWhenBatteryIsBelowLevel,
    [string]$Settings_File_Windows_Idle_Control_Keeper_txt,
    [string]$PycawAudioCheckerPath,
    [string]$PythonPath,
    [string]$DontSleepWhileThisFileExistsPath,
    [string]$MutualExclusionFlagFile,
    [string]$LogPath,
    [int]$LogMaxAgeDays,
    [int]$LogMaxSizeMB,
    [int]$LogToFileIntervalSeconds,
    [bool]$LogToConsoleVerbose,
    [int]$AbortWindowCountdownSeconds,
    [int]$CriticalBatteryAbortWindowCountdownSeconds,
    [int]$SettingsPollIntervalMinutes,
    [int]$FileSettingsPollIntervalMinutes,
    [int]$FailsafeTimeMinutes
)

# ───────────────────────────────────────────────────────────────────────────────
# 2. CONFIGURATION DEFAULTS (single source of truth)
# ───────────────────────────────────────────────────────────────────────────────
$script:Config = @{
	PauseScript 											= $false
	IgnoreTheSettingsFile									= $false
    PreventAndReplaceWindowsAutoSleep           			= $true
    FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan 	= $true
    UserSpecifiedSleepIdleTimeMinutes           			= 0
    UserSpecifiedHibernateIdleTimeMinutes       			= 0
    TurnOnScreensaverAtThisIdleTimeMinutes      			= 0
    TurnOffDisplayAtThisIdleTimeMinutes         			= 0
    CpuThresholdPercent                         			= 7
    GpuThresholdPercent                         			= 30
    DiskThresholdKBps                           			= 1250
    NetworkThresholdKBps                        			= 850
    ActiveSamplesWithinInterval                 			= 3
    ActivityDetectionPeriodSamples              			= 6
    ActivityDetectionPeriodSamplesAudio         			= 5
    LockPcAtThisIdleTimeMinutes                 			= 0
    IdleSecondsBeforeWeBroadcastSystemIdleEvent 			= 60
    UseOnlyInputAndAudioEventsForDisplayOff      			= $false
    DiskBlacklistDrives                         			= @("E", "F")
	SleepWhenBatteryIsBelowLevel							= 0
	HibernateWhenBatteryIsBelowLevel						= 0
    Settings_File_Windows_Idle_Control_Keeper_txt 			= "C:\Commands_And_Logs\[Settings_File]_Windows_Idle_Control_Keeper.txt"
    PycawAudioCheckerPath                       			= "C:\Commands_And_Logs\Pycaw_check_if_audio_is_playing.py"
    PythonPath                                  			= "$env:USERPROFILE\AppData\Local\Programs\Python\Python{VERSION}\python.exe"
    DontSleepWhileThisFileExistsPath                        = "D:\Bak\myLongSlowBackup.tmp"
    MutualExclusionFlagFile                        			= "C:\Commands_And_Logs\.WickMutualExclusionFlag"
    LogPath                                     			= "C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log"
    LogMaxAgeDays                               			= 30
    LogMaxSizeMB                                			= 10
    LogToFileIntervalSeconds                    			= 60
    LogToConsoleVerbose                         			= $true
    AbortWindowCountdownSeconds            					= 60
    CriticalBatteryAbortWindowCountdownSeconds            	= 60
    SettingsPollIntervalMinutes                 			= 1
    FileSettingsPollIntervalMinutes                 		= 1
    FailsafeTimeMinutes                         			= 0.98
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
		Write-Log "~~~~ Settings file not found at path: $Path."
		return 
	}
	
	Write-Log "||~*----- Checking config settings from file: $Path"
	if ($script:Config['IgnoreTheSettingsFile'] -eq $true) {
		Write-Log "Skipping loading most settings from file because it's set to be overwritten by the CLI param version, because -IgnoreTheSettingsFile:$true"
	} 
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
					
					#if ($script:Config['IgnoreTheSettingsFile'] -eq $true) {
					#	Write-Host "Skipping loading: $key ($val) from file because it's set to be overwritten by the CLI param version: $cliVersion"
					#}
					
					$val1Str = [string]$script:Config[$key]
					$val2Str = [string]$val
					
					if ($val1Str -ne $val2Str -and ($script:Config['IgnoreTheSettingsFile'] -eq $false -or $key -eq 'PauseScript')) {
						Write-Log "||||~*--- Updated config from file: '$key': from $($script:Config[$key]) to $val"
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
$script:g_DisplayTimeoutDurationMinutes = 0
$script:g_DisplayTurnedOff = $false
$script:g_ScreensaverTimeoutDurationMinutes = 0
$script:g_ScreenSaverStarted = $false
$script:g_PcLockedOnDemand = $false
$script:g_PreventSleep_ES = $false

$script:g_myUnixTimeEpochStart = Get-Date '2026-01-01'
$script:g_minutesPassedLastFrame = 0

$script:StartTime = Get-Date

$script:sleepOrHibernatePreventionFlagExists = $false

# --- Logging Setup ---
# logs to Windows > Event Viewer > Windows Logs > Application. It will have the date and time of the event. These can be queried by scripts.
function LogSystemEvent_IdleOn {
    [CmdletBinding()]
    param()
	
	$script:g_systemLoggedAsBeingIdle = $true

    $LogName   = "Application" # writing to the "System" log requires admin privileges
    $Source    = "Application" # writing to your new custom source e.g. "wick_idle_on" requires admin privileges
	$Category  = 69
    $EventId   = 420
    $EntryType = [System.Diagnostics.EventLogEntryType]::Warning # or Information
    $Message   = "[WICK: IDLE] System is idle according to the Windows_Idle_Control_Keeper.ps1. PID: $PID (IdleSecondsBeforeWeBroadcastSystemIdleEvent: $script:Config['IdleSecondsBeforeWeBroadcastSystemIdleEvent'].)"

    Write-EventLog -LogName $LogName -Source $Source -EventId $EventId -EntryType $EntryType -Message $Message -Category $Category
}

# logs to Windows > Event Viewer > Windows Logs > Application. It will have the date and time of the event. These can be queried by scripts.
function LogSystemEvent_IdleOff {
    [CmdletBinding()]
    param()

	$script:g_systemLoggedAsBeingIdle = $false

    $LogName   = "Application" # writing to the "System" log requires admin privileges
    $Source    = "Application" # writing to your new custom source e.g. "wick_idle_on" requires admin privileges
    $Category  = 69
    $EventId   = 421
    $EntryType = [System.Diagnostics.EventLogEntryType]::Warning # or Information
    $Message   = "[WICK: NOT Idle] System stopped being idle according to the Windows_Idle_Control_Keeper.ps1. PID: $PID"

    Write-EventLog -LogName $LogName -Source $Source -EventId $EventId -EntryType $EntryType -Message $Message -Category $Category
}

function LogSystemEvent_OnStart {
    [CmdletBinding()]
    param()
	
	$script:g_systemLoggedAsBeingIdle = $true

    $LogName   = "Application" # writing to the "System" log requires admin privileges
    $Source    = "Application" # writing to your new custom source e.g. "wick_on" requires admin privileges
	$Category  = 69
    $EventId   = 1111
    $EntryType = [System.Diagnostics.EventLogEntryType]::Warning # or Information
    $Message   = "[WICK: STARTED] The Windows_Idle_Control_Keeper.ps1 was started at $script:StartTime. PID: $PID"

    Write-EventLog -LogName $LogName -Source $Source -EventId $EventId -EntryType $EntryType -Message $Message -Category $Category
}

function LogSystemEvent_OnEnd {
    [CmdletBinding()]
    param()
	
	$script:g_systemLoggedAsBeingIdle = $true

    $LogName   = "Application" # writing to the "System" log requires admin privileges
    $Source    = "Application" # writing to your new custom source e.g. "wick_off" requires admin privileges
	$Category  = 69
    $EventId   = 1000
    $EntryType = [System.Diagnostics.EventLogEntryType]::Warning # or Information
    $Message   = "[WICK: ENDED] The Windows_Idle_Control_Keeper.ps1 that started at $script:StartTime has ended. PID: $PID"

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
    $logEntry = "[$timestamp][PID: $PID] [$Level] $Message"

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
	$battery = Get-WmiObject -Class BatteryStatus -Namespace root\wmi -ErrorAction SilentlyContinue
	if($battery) {
		return ($battery).PowerOnLine
	} else {
		return $true
	}
}

# returns 100 if not available. Note: you might want to also check if IsComputerPluggedIn but note that a laptop can discharge even while pllugged in in some cases.
function GetBatteryLevel{
	$battery = Get-WmiObject Win32_Battery -ErrorAction SilentlyContinue
	if($battery) {
		return ($battery).EstimatedChargeRemaining
	} else {
		return 100
	}
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
	if ($script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true -and $script:Config['LockPcAtThisIdleTimeMinutes'] -gt -1) {
		Write-Log "Locking PC because we're turning off the display and PreventAndReplaceWindowsAutoSleep is $true." "Info"
		Lock-PC
	}
	
	$script:g_DisplayTurnedOff = $true
	Write-Log "Turning off Display. g_DisplayTurnedOff: $script:g_DisplayTurnedOff" "Info"
	Write-Log " " "Info"
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
		Write-Log " " "Info"
	} else {
		Write-Log "Locked PC. g_PcLockedOnDemand: $script:g_PcLockedOnDemand" "INFO"
		Write-Log " " "Info"
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
	Write-Log " " "Info"
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
		$STANDBYIDLETIMOUT = "29f6c1db-86da-48c5-9fdb-f2b67b1f44da"
		
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
			$match = $regex.Match($hibernateOutput)
			
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
            #Write-Log "Power plan sleep idle timeout: disabled, using fallback: $($script:Config['UserSpecifiedSleepIdleTimeMinutes']) min" "INFO"
			$sleepMinutes = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
        } else {
			$sleepMinutes = [math]::Round($sleepSec / 60)
			#Write-Log "Power plan sleep idle timeout: $sleepMinutes min (sleep=$sleepSec sec)" "INFO"
		}
		
		# Convert to minutes, fallback if disabled
        if ($hibernateSec -le 0) {
            #Write-Log "Power plan hibernate idle timeout: disabled, using fallback: $($script:Config['UserSpecifiedHibernateIdleTimeMinutes']) min" "INFO"
            $hibernateMinutes = $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
        } else {
			$hibernateMinutes = [math]::Round($hibernateSec / 60)
			#Write-Log "Power plan hibernate idle timeout: $hibernateMinutes min (hibernate=$hibernateSec sec)" "INFO"
		}
				
		# Return named properties (PS 5.2 compatible)
		[PSCustomObject]@{
			sleepMinutesVal = $sleepMinutes
			hibernateMinutesVal = $hibernateMinutes
		}
    }
    catch {
        Write-Log "Failed to read power plan: $_, using fallback sleep: $($script:Config['UserSpecifiedSleepIdleTimeMinutes']) min, and fallback hibernate: $($script:Config['UserSpecifiedHibernateIdleTimeMinutes']) min" "WARN"
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

# --- Detect if session is unlocked, without admin ---
function Test-IsSessionUnlocked {
	# Get the active console session information
	$session = query session console 2>$null

	if ($session -match "Active") {
		return $true
	} else {
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
	Write-Log "Starting Sleep." "INFO"
	Write-Log " " "INFO"
    $result = [PowerManagement]::SetSuspendState($false, $true, $false)
    if (-not $result) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Log "Sleep failed! Win32 error: $err" "ERROR"
        Write-Log " " "ERROR"
    }
}

function Enter-HibernateState {
	Write-Log "Starting Hibernate." "INFO"
	Write-Log " " "INFO"
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
# Using the more accurate Processor Utility method: "Task Manager’s Processes and Performance tabs now use the “% Processor Utility” counter as the basis for their CPU numbers, rather than the “% Processor Time” counter Task Manager had relied upon and that is still used by Task Manager’s Details tab and by Sysinternals Process Explorer."
function Get-CpuViaProcessorUtility {
    <#
    .SYNOPSIS
        Returns the total system-wide CPU Utility percentage.
    .DESCRIPTION
        Uses the '\Processor Information(*)\% Processor Utility' counter.
        This provides the "modern" Task Manager percentage that accounts 
        for CPU frequency scaling/throttling.
    #>
    try {
        # 1. Query the specific 'Processor Information' counter
        # This counter is the one used by the Task Manager Performance tab.
        $counter = Get-Counter "\Processor Information(*)\% Processor Utility" -ErrorAction Stop

        # 2. The counter returns one sample per logical core.
        # To get the total system load, we must average the values of all cores.
        $coreValues = $counter.CounterSamples | ForEach-Object { $_.CookedValue }

        if ($null -eq $coreValues -or $coreValues.Count -eq 0) {
            return 0.0
        }

        # 3. Calculate the average across all cores
        # (Total Utility / Number of Cores) = System-wide Utility
        $systemWideUtility = ($coreValues | Measure-Object -Average).Average

        # We round to 2 decimal places for a clean number
        return [Math]::Round($systemWideUtility, 2)
    }
    catch {
        # If the counter is unavailable or fails, return 0 or a specific error value
        Write-Warning "Failed to retrieve CPU Utility: $_"
        return 0.0
    }
}

# --- GPUs ---
# call this every time you poll for new settings
function Update-GpuTypes {
	# detect if the gpu type exists
	$nv = [bool](Get-Command nvidia-smi -ErrorAction SilentlyContinue)
	$am = [bool](Get-Command rocm-smi -ErrorAction SilentlyContinue)
	if( $nv -ne $script:HasNvidia ){
		Write-Log "Nvidia gpu detected status changed: $($script:HasNvidia) -> $nv" "INFO"
		$script:HasNvidia = $nv
	}
	if( $am -ne $script:HasAmd ){
		Write-Log "AMD gpu detected status changed: $($script:HasAmd) -> $am" "INFO"
		$script:HasAmd = $am
	}
}
# If you have any number of AMD and/or Nvidia gpu(s), this function will get the maximum GPU usage percentage among all of those available gpus (so, a max, not an addition or an average or a median).
function Get-MaxGpuUsage {
    
    # detect if the gpu type exists
    if ($null -eq $script:HasNvidia -and $null -eq $script:HasAmd) {
        Update-GpuTypes
    }

    $allValues = @()

    # 2. NVIDIA Collection
    if ($script:HasNvidia) {
        # Using nounits returns a clean integer (e.g., "15")
        $nvOut = nvidia-smi --query-gpu=utilization.gpu --format=csv,noheader,nounits 2>$null
        if ($nvOut) {
            foreach ($line in $nvOut) {
                $cleanLine = $line.Trim()
                if ($cleanLine -match '^\d+$') {
                    $allValues += [int]$cleanLine
                }
            }
        }
    }

    # 3. AMD Collection (Using the CSV flag you discovered)
    if ($script:HasAmd) {
        # We query the CSV format. The columns are: device, GPU use (%), VRAM Total...
        # The index for 'GPU use (%)' is 1.
        $amdOut = rocm-smi --showuse --csv 2>$null
        if ($amdOut) {
            foreach ($line in $amdOut) {
                # Split the CSV line by comma
                $cols = $line.Split(',')
                # Check if the second column is a valid number
                if ($cols.Count -gt 1 -and $cols[1] -match '^\d+$') {
                    $allValues += [int]$cols[1]
                }
            }
        }
    }

    # 4. Final Calculation
    if ($allValues.Count -gt 0) {
        # Return the highest value found in the array
        return ($allValues | Measure-Object -Maximum).Maximum
    } else {
        # Return 0 if no GPUs were detected or no data was returned
        return 0
    }
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

        #$totalReadKBps = 0
        #$totalWriteKBps = 0
		$maxReadKBps = 0
        $maxWriteKBps = 0

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
			$readKBytes = $readBytes / 1KB
			$writeKBytes = $writeBytes / 1KB
            #$totalReadKBps += $readKBytes
            #$totalWriteKBps += $writeKBytes
			if($readKBytes -gt $maxReadKBps) {
				$maxReadKBps = $readKBytes
			}
			if($writeKBytes -gt $maxWriteKBps) {
				$maxWriteKBps = $writeKBytes
			}
        }

        #return [long]($totalReadKBps + $totalWriteKBps)
        return [long]($maxReadKBps + $maxWriteKBps)
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

# --- browse the latest python version you have ---
function Resolve-PythonPath {
    <#
    .SYNOPSIS
        Dynamically resolves Python path by scanning for the latest installed version.
    .DESCRIPTION
        Supports placeholder Python{VERSION} in config. Scans parent dir, 
        sorts directories numerically, and returns the path to the newest python.exe.
    #>
    [CmdletBinding()]
    param([string]$PathPattern)

    if ([string]::IsNullOrWhiteSpace($PathPattern)) { return $PathPattern }

    # Expand environment variables ($env:USERPROFILE -> C:\Users\...)
    $resolved = [Environment]::ExpandEnvironmentVariables($PathPattern)

    # Check if placeholder exists
    $placeholderIndex = $resolved.IndexOf('Python{VERSION}')
    if ($placeholderIndex -eq -1) { 
        Write-Verbose "No {VERSION} placeholder found. Using path as-is: $resolved"
        return $resolved 
    }

    # Extract base directory
    $basePath = $resolved.Substring(0, $placeholderIndex).TrimEnd('\')

    if (-not (Test-Path -LiteralPath $basePath)) {
        Write-Warning "Python base directory not found: $basePath"
        return $resolved
    }

    # Find all Python* folders
    $pyDirs = Get-ChildItem -LiteralPath $basePath -Directory -Filter 'Python*' -ErrorAction SilentlyContinue
    if ($pyDirs.Count -eq 0) {
        Write-Warning "No Python directories found in: $basePath"
        return $resolved
    }

    # 🔑 Sort numerically (Python313 > Python312 > Python39)
    $latest = $pyDirs | 
              Sort-Object { [long]($_.Name -replace '[^0-9]', '') } -Descending | 
              Select-Object -First 1

    # Build resolved path
    $finalPath = "$basePath\$($latest.Name)\python.exe"
    Write-Verbose "Resolved Python path: $finalPath"
    return $finalPath
}

$currentPythonPath = Resolve-PythonPath -PathPattern $script:Config['PythonPath']

# --- Sound: Detects if any audio is playing ---
function Get-AudioIsPlaying {
    if (-not [System.Environment]::UserInteractive) { return $false }

    # Capture stdout only, force case-insensitive comparison, guarantee boolean
	$pythonExe = $currentPythonPath #$script:Config['PythonPath']
	$pyScript  = $script:Config['PycawAudioCheckerPath']
	$rawOutput = & "$pythonExe" "$pyScript" 2>$null | ForEach-Object { $_.Trim().ToLower() }
	$isPlaying = $rawOutput -eq 'true'
	#Write-Log "isPlaying: $isPlaying" "INFO"
	return $isPlaying
}

# --- Show abort dialog (interactive only) ---
function Show-AbortDialog {
    [CmdletBinding()]
    param(
        [int]$Seconds,
        [string]$ActionName,
		[string]$Message
    )

    if (-not (Test-IsInteractiveSession)) {
        Write-Log "Non-interactive session - skipping message box." "DEBUG"
        return $false
    }

    # Ensure timeout is valid (0 means wait forever)
    if ($Seconds -le 0) { $Seconds = 10 }

    $title = "System $actionName Warning"
    $msg = "$Message`n`n" +
           "System will $actionName in $Seconds seconds.`n`n" +
           "Abort?"

    $wsh = New-Object -ComObject WScript.Shell
    # Cast to [int] to guarantee COM respects the timeout
    $result = $wsh.Popup($msg, [int]$Seconds, $title, 36) # 36 = vbYesNo + vbQuestion

    # Return $true ONLY if user pressed Yes (abort)
    # 6 = Yes, 7 = No, 0/-1 = Timeout (all proceed = $false)
    return $result -eq 6
}

# --- Makes sure only one instance of this script runs at a time. Read the MutualExclusionFlagFile param description for why. ---

function CreateOrUpdate-MutualExclusionFlagFile {
    <#
    .SYNOPSIS
        Atomically overwrites the mutual exclusion flag with the current process PID.
    #>
    [CmdletBinding()]
    param()

    $filePath = $script:Config['MutualExclusionFlagFile']
    $parentDir = Split-Path -Path $filePath -Parent

    if (-not (Test-Path -LiteralPath $parentDir)) {
		Write-Log "[CreateOrUpdate-MutualExclusionFlagFile] ERROR: Invalid directory for MutualExclusionFlagFile. Multiple instances may run concurrently. FilePath: $filePath" "ERROR"
        return
    }

    # Atomic overwrite: completely replaces file contents with current PID
    [System.IO.File]::WriteAllText($filePath, $PID.ToString(), [System.Text.Encoding]::UTF8)
}

# comparing file contents to $script:StartTime + $script:deltaTimeOfFlagTouch
function WeOwn-MutualExclusionFlagFile {
    <#
    .SYNOPSIS
        Checks if this script instance currently "owns" the mutual exclusion flag.
    .DESCRIPTION
        Returns $true if the PID in the file matches our PID.
        Returns $false if another instance's PID is in the file.
    #>
    [CmdletBinding()]
    param()

    $filePath = $script:Config['MutualExclusionFlagFile']
    
    if (-not (Test-Path -LiteralPath $filePath)) { 
		Write-Log "[WeOwn-MutualExclusionFlagFile] no file found at path: $filePath" "ERROR"
		return $true 
	}

    try {
        $filePID = [int]::Parse([System.IO.File]::ReadAllText($filePath).Trim())
        # If the file contains OUR PID, we own the flag and can proceed
        return $filePID -eq $PID
    }
    catch {
		Write-Log "[WeOwn-MutualExclusionFlagFile] Failed to parse MutualExclusionFlagFile: $_" "ERROR"
        return $true
    }
}

# Test-Path-Timeouted -timeoutMilliseconds 6000 -pathToTest "\\Network_Paths\On_Windows\offline_network_paths\can\take\22s\to\a\minute\before_they_return_false"
function Test-Path-Timeouted-Scope {
	[CmdletBinding()]
    param(
	    [int]$timeoutMilliseconds,
        [string]$pathToTest
    )
	
	if($timeoutMilliseconds -le 1) {
		Write-Log "[Test-Path-Timeouted] You didn't provide a timeoutMilliseconds parameter: $timeoutMilliseconds for this path: $pathToTest" "WARN"
		$timeoutMilliseconds = 5000
	}
	
	$ps = [powershell]::Create().AddScript("Test-Path '$pathToTest'")
	$handle = $ps.BeginInvoke()
	if(-not $handle.AsyncWaitHandle.WaitOne($timeoutMilliseconds)){
		$boolExists = $false
	} else {
		$boolExists = $true
	}
	$boolExists
	if ($boolExists) {
		$result = $ps.EndInvoke($handle)
	} else {
		$result = $boolExists
	}
	#Write-Log "[Test-Path-Timeouted] ---------- boolExists: $boolExists, result: $result" "WARN"
	return $result -and $boolExists
}

function Test-Path-Timeouted {
	[CmdletBinding()]
    param(
	    [int]$timeoutMilliseconds,
        [string]$pathToTest
    )
	
	$res0 = Test-Path-Timeouted-Scope -timeoutMilliseconds $timeoutMilliseconds -pathToTest $pathToTest
	$res = "s: " + $res0
	if($res -ceq "s: True True") {
		#Write-Log "[Test-Path-Timeouted] ----------r true res: $res" "WARN"
		#return $true
		$script:sleepOrHibernatePreventionFlagExists = $true
	} else {
		#Write-Log "[Test-Path-Timeouted] ----------r false res: $res" "WARN"
		#return $false
		$script:sleepOrHibernatePreventionFlagExists = $false
	}
}

# --- Main Sequence ---
LogSystemEvent_OnStart
CreateOrUpdate-MutualExclusionFlagFile
Write-Log " " "INFO"
Write-Log "~*------- W.I.C.K. started. ---------" "INFO"
Write-Log "~*------- PID: $PID ---------" "INFO"
Write-Log "Log path: $($script:Config['LogPath'])" "INFO"
Write-Log "Settings path: $($script:Config['Settings_File_Windows_Idle_Control_Keeper_txt'])" "INFO"

if ($script:Config['FileSettingsPollIntervalMinutes'] -ne 0.0) {
	Update-ConfigFromSettingsFile
}

$script:g_isPluggedIn = IsComputerPluggedIn
$script:g_batteryLevel = GetBatteryLevel
Write-Log "System plugged in: $script:g_isPluggedIn, system battery level: $script:g_batteryLevel" "INFO"

#Write-Log "  Dynamic idle timeout (checking the current active power plan value every: $($script:Config['SettingsPollIntervalMinutes']))"
if ($script:Config['FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan']) {
	$tuple = Get-PowerPlanIdleTimeoutMinutes $script:g_isPluggedIn
	$script:g_CurrentSleepIdleTimeMinutes = $tuple.sleepMinutesVal
	$script:g_CurrentHibernateIdleTimeMinutes = $tuple.hibernateMinutesVal
    Write-Log "You set to use windows power plan's sleep and hiberante values (if sleep/hibernate are enabled on the system): $($script:g_CurrentSleepIdleTimeMinutes) min and $($script:g_CurrentHibernateIdleTimeMinutes) min. (We check to update this value every: SettingsPollIntervalMinutes: $($script:Config['SettingsPollIntervalMinutes']) min.)" "INFO"
	
	if($script:Config['UserSpecifiedSleepIdleTimeMinutes'] -gt 0) {
		Write-Log "    BUT, you specified UserSpecifiedSleepIdleTimeMinutes: $($script:Config['UserSpecifiedSleepIdleTimeMinutes']), so we override the $script:g_CurrentSleepIdleTimeMinutes min to $($script:Config['UserSpecifiedSleepIdleTimeMinutes']) min." "INFO"
		$script:g_CurrentSleepIdleTimeMinutes = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
	}
	
	if($script:Config['UserSpecifiedHibernateIdleTimeMinutes'] -gt 0) {
		Write-Log "    BUT, you specified UserSpecifiedHibernateIdleTimeMinutes: $($script:Config['UserSpecifiedHibernateIdleTimeMinutes']), so we override the $script:g_CurrentHibernateIdleTimeMinutes min to $($script:Config['UserSpecifiedHibernateIdleTimeMinutes']) min." "INFO"
		$script:g_CurrentHibernateIdleTimeMinutes = $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
	}
} else {
	$script:g_CurrentSleepIdleTimeMinutes = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
	$script:g_CurrentHibernateIdleTimeMinutes = $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
    Write-Log "Using manual sleep and hibernate timeout values (if sleep/hibernate are enabled on the system): $($script:Config['UserSpecifiedSleepIdleTimeMinutes']) min and $($script:Config['UserSpecifiedHibernateIdleTimeMinutes']) min." "INFO"
}

if ($script:Config['PreventAndReplaceWindowsAutoSleep']) {
	Write-Log "Because PreventAndReplaceWindowsAutoSleep: $($script:Config['PreventAndReplaceWindowsAutoSleep']), we need to manually trigger the display to turn off at the power plan's display setting time (and also lock), and also trigger the screensaver at its time if it exists. Values below." "INFO"
		
	if ($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'] -gt 0.0) {
		$script:g_DisplayTimeoutDurationMinutes = $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']
		Write-Log "Using user set display timeout: $($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'])" "INFO"
	} else {
		$script:g_DisplayTimeoutDurationMinutes = $(Get-DisplayTimeoutSeconds $script:g_isPluggedIn) / 60
		Write-Log "Using windows power plan's g_DisplayTimeoutDurationMinutes: $($script:g_DisplayTimeoutDurationMinutes). We need this because Laptops will stop auto turning off their display if you tell them to not sleep (by using SetThreadExecutionState ES_SYSTEM_REQUIRED). (We check to update this value every: SettingsPollIntervalMinutes: $($script:Config['SettingsPollIntervalMinutes']) min.)" "INFO"
	}
	
	if($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'] -gt 0.0) {
		$script:g_ScreensaverTimeoutDurationMinutes = $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']
		Write-Log "Using user set screensaver timeout: $($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'])" "INFO"
	} else {
		$script:g_ScreensaverTimeoutDurationMinutes = $(Get-ScreensaverTimeoutSeconds) / 60
		Write-Log "Using windows power plan's g_ScreensaverTimeoutDurationMinutes: $($script:g_ScreensaverTimeoutDurationMinutes) (in case you use the screensaver). We need this because Laptops will stop auto turning on the screensaver if you tell them to not sleep (by using SetThreadExecutionState ES_SYSTEM_REQUIRED). (We check to update this value every: SettingsPollIntervalMinutes: $($script:Config['SettingsPollIntervalMinutes']) min.)" "INFO"
	}
} elseif ($script:Config['FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan'] -eq $false) {
	if ($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'] -gt 0.0) {
		$script:g_DisplayTimeoutDurationMinutes = $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']
	}
	
	if($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'] -gt 0.0) {
		$script:g_ScreensaverTimeoutDurationMinutes = $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']
	}
	Write-Log "Display-off timeout and Screensaver timeout (if available): $script:g_DisplayTimeoutDurationMinutes min and $script:g_ScreensaverTimeoutDurationMinutes min." "INFO"
} else {
	if ($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'] -gt 0.0) {
		$script:g_DisplayTimeoutDurationMinutes = $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']
	} else {
		$script:g_DisplayTimeoutDurationMinutes = $(Get-DisplayTimeoutSeconds $script:g_isPluggedIn) / 60
	}
	if($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'] -gt 0.0) {
		$script:g_ScreensaverTimeoutDurationMinutes = $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']
	} else {
		$script:g_ScreensaverTimeoutDurationMinutes = $(Get-ScreensaverTimeoutSeconds) / 60
	}
	Write-Log "Display-off timeout and Screensaver timeout (if available): $script:g_DisplayTimeoutDurationMinutes min and $script:g_ScreensaverTimeoutDurationMinutes min." "INFO"
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
$script:g_maxSamples = [int]([math]::Ceiling($script:Config['ActivityDetectionPeriodSamples']))# / $script:Config['SampleIntervalSec']))
$script:g_maxSamplesAudio = [int]([math]::Ceiling($script:Config['ActivityDetectionPeriodSamplesAudio']))# / $script:Config['SampleIntervalSec']))
$script:g_idleSeconds = 0.0
$script:g_idleSeconds_userOrAudioActivity = 0.0
$script:g_systemLoggedAsBeingIdle = $false
$script:g_sw = [System.Diagnostics.Stopwatch]::StartNew()
$script:g_lastElapsedSeconds = 0.0

$script:g_cpuHistory = New-Object 'System.Collections.Queue' $script:g_maxSamples
$script:g_gpuHistory = New-Object 'System.Collections.Queue' $script:g_maxSamples
$script:g_diskHistory = New-Object 'System.Collections.Queue' $script:g_maxSamples
$script:g_netHistory = New-Object 'System.Collections.Queue' $script:g_maxSamples
$script:g_audioHistory = New-Object 'System.Collections.Queue' $script:g_maxSamplesAudio

$script:g_nextSettingsPollSeconds = $script:Config['SettingsPollIntervalMinutes'] * 60
$script:g_nextFileSettingsPollSeconds = $script:Config['FileSettingsPollIntervalMinutes'] * 60
# $script:sleepOrHibernatePreventionFlagExists = Test-Path $script:Config['DontSleepWhileThisFileExistsPath']
Test-Path-Timeouted -timeoutMilliseconds 6000 -pathToTest $script:Config['DontSleepWhileThisFileExistsPath']
$flagFile = $script:Config['DontSleepWhileThisFileExistsPath']
if($script:sleepOrHibernatePreventionFlagExists -eq $true) {
	Write-Log "We won't sleep or hibernate, while the DontSleepWhileThisFileExistsPath flag file or folder exists: ($flagFile)" "INFO"
}

if ($script:Config['PauseScript'] -eq $true) {
	Write-Log "Script pause flag is present - while loop running but skipping until flag removed or renamed." "INFO"
}

# NOTE: if this script is frozen or closed unexpectedly in a way that [WindowsSleepWrangler]::StopIgnoringIdleTimers() doesn't get called, windows may not go to sleep again until it is called or until it's restarted. But since we use try - finally, it should auto clean up after itself unless you somehow freeze the thread.
if($script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true){
	$script:g_PreventSleep_ES = $true
	[WindowsSleepWrangler]::IgnoreIdleTimers()
}

try {
	# --- Main Loop ---
	while ($true) {
		$skipThisFrame = $false
		#Write-Log "Tick. idleSeconds: $script:g_idleSeconds" "INFO"
		#Write-Host-Wrapper "Tick. idleSeconds: $script:g_idleSeconds" "INFO"
		
		$currentElapsedSeconds = $script:g_sw.Elapsed.TotalSeconds
		$deltaTimeSeconds = $currentElapsedSeconds - $script:g_lastElapsedSeconds
		$script:g_lastElapsedSeconds = $currentElapsedSeconds
		
		$minutesPassed = (Get-Date).Subtract($script:g_myUnixTimeEpochStart).TotalMinutes
		# check if it's been more than one minute since the script updated -- it means we woke up from sleep
		$updateDiffInMinutes = $minutesPassed - $script:g_minutesPassedLastFrame
		if($updateDiffInMinutes -gt 1) {
			$script:g_idleSeconds = -1.0 * $script:Config['FailsafeTimeMinutes'] * 60 # we should set it to 0 here but I do a failsafe time here in case somebody screws something up / adds something that for example would lock the pc every second. This way if you sleep + wake, or restart the pc, you get 60 seconds to stop it even if you set it to run hidden on system startup from task scheduler.
			$script:g_idleSeconds_userOrAudioActivity = -1.0 * $script:Config['FailsafeTimeMinutes'] * 60
			
			$deltaTimeSeconds = 0.0
			Write-Log "It's been $updateDiffInMinutes minute(s) since the last update, which means we just started, or were sleeping, or somehow lagging a lot. Resetting idle counter, with failsafe, to: $script:g_idleSeconds." "INFO"
		}
		$script:g_minutesPassedLastFrame = $minutesPassed
		
		if ($script:g_idleSeconds -lt 0) {
			$script:g_idleSeconds += $deltaTimeSeconds
			
			if ($script:g_idleSeconds -gt (-1.0 * $script:Config['FailsafeTimeMinutes'] * 60 + 5) -and $script:g_idleSeconds -lt -5) {
				Write-Host-Wrapper "Failsafe period: $script:g_idleSeconds s < 0. [deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
			} else {
				Write-Log "Failsafe period: $script:g_idleSeconds s < 0. [deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
			}
			$skipThisFrame = $true
		}
		
		if ($script:g_idleSeconds -gt $script:Config['IdleSecondsBeforeWeBroadcastSystemIdleEvent'] -and $script:g_systemLoggedAsBeingIdle -eq $false) {
			LogSystemEvent_IdleOn
		}
		
		$script:g_nextFileSettingsPollSeconds = $script:g_nextFileSettingsPollSeconds - $deltaTimeSeconds
		$checkedSettings = $false
		if ($script:g_nextFileSettingsPollSeconds -le 0 -and $script:Config['FileSettingsPollIntervalMinutes'] -ne 0.0) {
			Update-ConfigFromSettingsFile
			$checkedSettings = $true
			$script:g_nextFileSettingsPollSeconds = $script:Config['FileSettingsPollIntervalMinutes'] * 60
			
			$flagFile = $script:Config['DontSleepWhileThisFileExistsPath']
			#$dsfexists = Test-Path $flagFile
			$dsfexistsOld = $script:sleepOrHibernatePreventionFlagExists
			Test-Path-Timeouted -timeoutMilliseconds 6000 -pathToTest $flagFile
			if ($script:sleepOrHibernatePreventionFlagExists -ne $dsfexistsOld) {
				Write-Log "DontSleepWhileThisFileExistsPath file flag ($flagFile) went from $dsfexistsOld to $script:sleepOrHibernatePreventionFlagExists" "INFO"
			}
			if($script:sleepOrHibernatePreventionFlagExists -eq $true) {
				Write-Log "We won't sleep or hibernate, because of this prevention file or folder $flagFile" "INFO"
			}
		}
		
		# Poll power plan timeout every $script:Config['SettingsPollIntervalMinutes'] updates
		$script:g_nextSettingsPollSeconds = $script:g_nextSettingsPollSeconds - $deltaTimeSeconds
		if ($script:g_nextSettingsPollSeconds -le 0) {
			# exit this script if a new script was started on this computer.
			$weOwnMuttex = WeOwn-MutualExclusionFlagFile
			if($weOwnMuttex -eq $false) {
				Write-Log "~~~~~~~~~ Exiting this WICK script we started at $script:StartTime PID: $PID because somebody else (hopefully another WICK script) updated the PID in the MutualExclusionFlagFile since then). ~~~~~~~*-" "INFO"
				break;
			}
			
			# Sliding windows for sustained detection
			if($script:g_maxSamples -ne $script:Config['ActivityDetectionPeriodSamples']) {
				Write-Log "ActivityDetectionPeriodSamples changed: $script:g_maxSamples to $($script:Config['ActivityDetectionPeriodSamples'])." "INFO"
				$script:g_maxSamples = $script:Config['ActivityDetectionPeriodSamples']
			}
			if($script:g_maxSamplesAudio -ne $script:Config['ActivityDetectionPeriodSamplesAudio']) {
				Write-Log "ActivityDetectionPeriodSamplesAudio changed: $script:g_maxSamplesAudio to $($script:Config['ActivityDetectionPeriodSamplesAudio'])." "INFO"
				$script:g_maxSamplesAudio = $script:Config['ActivityDetectionPeriodSamplesAudio']
			}
			
			$script:g_isPluggedIn = IsComputerPluggedIn
			if ($script:Config['FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan']) {
				$tuple = Get-PowerPlanIdleTimeoutMinutes $script:g_isPluggedIn
				$newSleepTimeout = $tuple.sleepMinutesVal
				$newHibernateTimeout = $tuple.hibernateMinutesVal
				
				if($script:Config['UserSpecifiedSleepIdleTimeMinutes'] -gt 0.0){
					if($script:g_CurrentSleepIdleTimeMinutes -ne $script:Config['UserSpecifiedSleepIdleTimeMinutes']) {
						Write-Log "You specified UserSpecifiedSleepIdleTimeMinutes: $($script:Config['UserSpecifiedSleepIdleTimeMinutes']), so we override the power plan's value of $newSleepTimeout min." "INFO"
						$script:g_CurrentSleepIdleTimeMinutes = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
					}
				} elseif ($newSleepTimeout -ne $script:g_CurrentSleepIdleTimeMinutes) {
					Write-Log "Power plan sleep timeout changed: $script:g_CurrentSleepIdleTimeMinutes to $newSleepTimeout min." "INFO"
					$script:g_CurrentSleepIdleTimeMinutes = $newSleepTimeout
				}
				
				if($script:Config['UserSpecifiedHibernateIdleTimeMinutes'] -gt 0.0) {
					if($script:g_CurrentHibernateIdleTimeMinutes -ne $script:Config['UserSpecifiedHibernateIdleTimeMinutes']) {
						Write-Log "You specified UserSpecifiedHibernateIdleTimeMinutes: $($script:Config['UserSpecifiedHibernateIdleTimeMinutes']), so we override the power plan's value of $newSleepTimeout min." "INFO"
						$script:g_CurrentHibernateIdleTimeMinutes = $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
					}
				} elseif ($newHibernateTimeout -ne $script:g_CurrentHibernateIdleTimeMinutes) {
					Write-Log "Power plan hibernate timeout changed: $script:g_CurrentHibernateIdleTimeMinutes to $newHibernateTimeout min." "INFO"
					$script:g_CurrentHibernateIdleTimeMinutes = $newHibernateTimeout
				}
			} elseif ($script:g_CurrentSleepIdleTimeMinutes -ne $script:Config['UserSpecifiedSleepIdleTimeMinutes'] -or $script:g_CurrentHibernateIdleTimeMinutes -ne $script:Config['UserSpecifiedHibernateIdleTimeMinutes']) {
				$script:g_CurrentSleepIdleTimeMinutes = $script:Config['UserSpecifiedSleepIdleTimeMinutes']
				$script:g_CurrentHibernateIdleTimeMinutes = $script:Config['UserSpecifiedHibernateIdleTimeMinutes']
				Write-Log "Updating manual sleep and hibernate timeout values (if sleep/hibernate are enabled on the system): $($script:Config['UserSpecifiedSleepIdleTimeMinutes']) min and $($script:Config['UserSpecifiedHibernateIdleTimeMinutes']) min." "INFO"
			}
			
			if ($script:Config['PreventAndReplaceWindowsAutoSleep']) {
				if ($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'] -gt 0.0) {
					if($script:g_DisplayTimeoutDurationMinutes -ne $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']) {
						$script:g_DisplayTimeoutDurationMinutes = $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']
						Write-Log "Updating user set display timeout: $($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'])" "INFO"
					}
				} else {
					$newDisplayTimeoutDurationMinutes = $(Get-DisplayTimeoutSeconds $script:g_isPluggedIn) / 60
					if($newDisplayTimeoutDurationMinutes -ne $script:g_DisplayTimeoutDurationMinutes) {
						Write-Log "System g_DisplayTimeoutDurationMinutes changed: $script:g_DisplayTimeoutDurationMinutes to $newDisplayTimeoutDurationMinutes minutes." "INFO"
						$script:g_DisplayTimeoutDurationMinutes = $newDisplayTimeoutDurationMinutes
					}
				}
				
				if($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'] -gt 0.0) {
					if($script:g_ScreensaverTimeoutDurationMinutes -ne $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']) {
						$script:g_ScreensaverTimeoutDurationMinutes = $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']
						Write-Log "Updating user set screensaver timeout: $($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'])" "INFO"
					}
				} else {
					$newScreensaverTimeoutDurationMinutes = $(Get-ScreensaverTimeoutSeconds) / 60
					if($newScreensaverTimeoutDurationMinutes -ne $script:g_ScreensaverTimeoutDurationMinutes) {
						Write-Log "System g_ScreensaverTimeoutDurationMinutes changed: $script:g_ScreensaverTimeoutDurationMinutes to $newScreensaverTimeoutDurationMinutes seconds." "INFO"
						$script:g_ScreensaverTimeoutDurationMinutes = $newScreensaverTimeoutDurationMinutes
					}
				}
			} elseif ($script:Config['FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan'] -eq $false) {
				if ($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'] -gt 0.0) {
					if ($script:g_DisplayTimeoutDurationMinutes -ne $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']) {
						$script:g_DisplayTimeoutDurationMinutes = $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']
						Write-Log "Display-off timeout (if available): $script:g_DisplayTimeoutDurationMinutes min." "INFO"
					}
				}
				
				if($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'] -gt 0.0) {
					if ($script:g_ScreensaverTimeoutDurationMinutes -ne $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']) {
						$script:g_ScreensaverTimeoutDurationMinutes = $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']
						Write-Log "Screensaver timeout (if available): $script:g_ScreensaverTimeoutDurationMinutes min." "INFO"
					}
				}
			} else {
				if ($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'] -gt 0.0) {
					if ($script:g_DisplayTimeoutDurationMinutes -ne $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']) {
						$script:g_DisplayTimeoutDurationMinutes = $script:Config['TurnOffDisplayAtThisIdleTimeMinutes']
						Write-Log "Display-off timeout (if available): $script:g_DisplayTimeoutDurationMinutes min." "INFO"
					}
				} else {
					$displayTimeout = $(Get-DisplayTimeoutSeconds $script:g_isPluggedIn) / 60
					if($script:g_DisplayTimeoutDurationMinutes -ne $displayTimeout) {
						$script:g_DisplayTimeoutDurationMinutes = $displayTimeout
						Write-Log "Display-off timeout (if available): $script:g_DisplayTimeoutDurationMinutes min." "INFO"
					}
				}
				
				if($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'] -gt 0.0) {
					if ($script:g_ScreensaverTimeoutDurationMinutes -ne $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']) {
						$script:g_ScreensaverTimeoutDurationMinutes = $script:Config['TurnOnScreensaverAtThisIdleTimeMinutes']
						Write-Log "Screensaver timeout (if available): $script:g_ScreensaverTimeoutDurationMinutes min." "INFO"
					}
				} else {
					$screensaverTimeout = $(Get-ScreensaverTimeoutSeconds) / 60
					if ($script:g_ScreensaverTimeoutDurationMinutes -ne $screensaverTimeout) {
						$script:g_ScreensaverTimeoutDurationMinutes = $screensaverTimeout
						Write-Log "Screensaver timeout (if available): $script:g_ScreensaverTimeoutDurationMinutes min." "INFO"
					}
				}
			}
			
			# check maybe the user updated python in the meantime (this script is meant to run pretty much constantly)
			$currentPythonPath = Resolve-PythonPath -PathPattern $script:Config['PythonPath']
			
			Update-GpuTypes
			
			$script:g_nextSettingsPollSeconds = $script:Config['SettingsPollIntervalMinutes'] * 60
		}
		
		if ($script:Config['PauseScript'] -eq $true) {
			$script:g_idleSeconds = 0.0
			$script:g_idleSeconds_userOrAudioActivity = 0.0
			if ($script:g_systemLoggedAsBeingIdle -eq $true) {
				LogSystemEvent_IdleOff
			}
			if ($script:g_PreventSleep_ES -eq $true -and $script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true) {
				$script:g_PreventSleep_ES = $false
				[WindowsSleepWrangler]::StopIgnoringIdleTimers()
			}
			$skipThisFrame = $true
		} elseif ($script:g_PreventSleep_ES -eq $false -and $script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true) {
			$script:g_PreventSleep_ES = $true
			[WindowsSleepWrangler]::IgnoreIdleTimers()
		}
		
		if($skipThisFrame -eq $false) {
			
			# CPU
			$cpu = Get-CpuViaProcessorUtility
			$cpuAbove = $cpu -gt $script:Config['CpuThresholdPercent']
			$script:g_cpuHistory.Enqueue($cpuAbove)
			if ($script:g_cpuHistory.Count -gt $script:g_maxSamples) { 
				$null = $script:g_cpuHistory.Dequeue()
			}
			
			# GPU
			$gpu = Get-MaxGpuUsage
			$gpuAbove = $gpu -gt $script:Config['GpuThresholdPercent']
			$script:g_gpuHistory.Enqueue($gpuAbove)
			if ($script:g_gpuHistory.Count -gt $script:g_maxSamples) { 
				$null = $script:g_gpuHistory.Dequeue()
			}

			# Disk
			$disk = Get-DiskIoKBps
			$diskAbove = $disk -gt $script:Config['DiskThresholdKBps']
			#if($diskAbove -eq $true) {
				# Play a sound to debug when an array of write events happen - so you can quickly look at process explorer or something.
				#(New-Object System.Media.SoundPlayer "$env:windir\Media\Ring06.wav").Play()
				#Write-Log "|||||| DISK ABOVE THRESHOLD: $diskAbove ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||" "INFO"
			#}
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

			# Mouse, replaced with Get-SecondsSinceLastInputInfo
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
						Write-Log " " "INFO"
						Write-Log "[IDLE BREAKER] CPU: $activeCount/$script:g_maxSamples samples > $($script:Config['CpuThresholdPercent'])% (>= $($script:Config['ActiveSamplesWithinInterval']) required). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Log " " "INFO"
					} elseif ($script:Config['LogToConsoleVerbose']) {
						Write-Host-Wrapper " " "INFO"
						Write-Host-Wrapper "[IDLE BREAKER] CPU: $activeCount/$script:g_maxSamples samples > $($script:Config['CpuThresholdPercent'])% (>= $($script:Config['ActiveSamplesWithinInterval']) required). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Host-Wrapper " " "INFO"
					}
					$hasSustainedActivity = $true
				}
			}
			
			# GPU: >= $script:Config['ActiveSamplesWithinInterval'] of last $script:g_maxSamples were above threshold
			if ($script:g_gpuHistory.Count -ge $script:g_maxSamples) {
				$activeCount = ($script:g_gpuHistory | Where-Object { $_ }).Count
				if ($activeCount -ge $script:Config['ActiveSamplesWithinInterval']) {
					if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
						Write-Log " " "INFO"
						Write-Log "[IDLE BREAKER] GPU: $activeCount/$script:g_maxSamples samples > $($script:Config['GpuThresholdPercent'])% (>= $($script:Config['ActiveSamplesWithinInterval']) required). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Log " " "INFO"
					} elseif ($script:Config['LogToConsoleVerbose']) {
						Write-Host-Wrapper " " "INFO"
						Write-Host-Wrapper "[IDLE BREAKER] GPU: $activeCount/$script:g_maxSamples samples > $($script:Config['GpuThresholdPercent'])% (>= $($script:Config['ActiveSamplesWithinInterval']) required). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Host-Wrapper " " "INFO"
					}
					$hasSustainedActivity = $true
				}
			}

			# Disk: >= $script:Config['ActiveSamplesWithinInterval'] of last $script:g_maxSamples were above threshold
			if ($script:g_diskHistory.Count -ge $script:g_maxSamples) {
				$activeCount = ($script:g_diskHistory | Where-Object { $_ }).Count
				if ($activeCount -ge $script:Config['ActiveSamplesWithinInterval']) {
					if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
						Write-Log " " "INFO"
						Write-Log "[IDLE BREAKER] Disk: $activeCount/$script:g_maxSamples samples > $($script:Config['DiskThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required).[idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Log " " "INFO"
					} elseif ($script:Config['LogToConsoleVerbose']) {
						Write-Host-Wrapper " " "INFO"
						Write-Host-Wrapper "[IDLE BREAKER] Disk: $activeCount/$script:g_maxSamples samples > $($script:Config['DiskThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Host-Wrapper " " "INFO"
					}
					$hasSustainedActivity = $true
				}
				
			}

			# Network: >= $script:Config['ActiveSamplesWithinInterval'] of last $script:g_maxSamples were above threshold
			if ($script:g_netHistory.Count -ge $script:g_maxSamples) {
				$activeCount = ($script:g_netHistory | Where-Object { $_ }).Count
				if ($activeCount -ge $script:Config['ActiveSamplesWithinInterval']) {
					if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
						Write-Log " " "INFO"
						Write-Log "[IDLE BREAKER] Network: $activeCount/$script:g_maxSamples samples > $($script:Config['NetworkThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Log " " "INFO"
					} elseif ($script:Config['LogToConsoleVerbose']) {
						Write-Host-Wrapper " " "INFO"
						Write-Host-Wrapper "[IDLE BREAKER] Network: $activeCount/$script:g_maxSamples samples > $($script:Config['NetworkThresholdKBps']) KBps (>= $($script:Config['ActiveSamplesWithinInterval']) required). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Host-Wrapper " " "INFO"
					}
					$hasSustainedActivity = $true
				}
			}

			$audioBasedActivityThisFrame = $false
			$inputBasedActivityThisFrame = $false

			# Audio: all samples must be true
			if ($script:g_audioHistory.Count -eq $script:g_maxSamplesAudio -and ($script:g_audioHistory | Where-Object { $_ }).Count -eq $script:g_maxSamplesAudio) {
				if($script:g_idleSeconds -ge $script:Config['ActivityDetectionPeriodSamplesAudio']){
					if ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds']) {
						Write-Log " " "INFO"
						Write-Log "[IDLE BREAKER] Sustained audio playing for $($script:Config['ActivityDetectionPeriodSamplesAudio']) samples, Resetting idle counter. [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Log " " "INFO"
					} elseif ($script:Config['LogToConsoleVerbose']) {
						Write-Host-Wrapper " " "INFO"
						Write-Host-Wrapper "[IDLE BREAKER] Sustained audio playing for $($script:Config['ActivityDetectionPeriodSamplesAudio']) samples, Resetting idle counter. [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Host-Wrapper " " "INFO"
					}
					$hasSustainedActivity = $true
					$audioBasedActivityThisFrame = $true
				}
			}

			$secondsSinceLastInputInfo = Get-SecondsSinceLastInputInfo
			$idleSecOrUserSec = $script:g_idleSeconds
			if ($script:Config['UseOnlyInputAndAudioEventsForDisplayOff'] -eq $true) {
				$idleSecOrUserSec = $script:g_idleSeconds_userOrAudioActivity
			}
			if ($secondsSinceLastInputInfo -le $idleSecOrUserSec) {
				#if ($script:g_idleSeconds -ge $script:Config['ActivityDetectionPeriodSamples']) {
					if($idleSecOrUserSec -ge $script:Config['LogToFileIntervalSeconds']){
						Write-Log " " "INFO"
						#Write-Log "[IDLE BREAKER][idleSeconds: $([math]::Round($script:g_idleSeconds, 5))] Mouse moved $([math]::Round($mouseDelta,1)) px > $MouseThresholdPixels, Resetting idle counter." "INFO"
						Write-Log "[IDLE BREAKER] Mouse/touch/keyboard activity registered $secondsSinceLastInputInfo seconds ago. Resetting idle counter. [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Log " " "INFO"
					} elseif ($script:Config['LogToConsoleVerbose'] -eq $true) {
						Write-Host-Wrapper " " "INFO"
						#Write-Host-Wrapper "[IDLE BREAKER][idleSeconds: $([math]::Round($script:g_idleSeconds, 5))] Mouse moved $([math]::Round($mouseDelta,1)) px > $MouseThresholdPixels, Resetting idle counter." "INFO"
						Write-Host-Wrapper "[IDLE BREAKER] Mouse/touch/keyboard activity registered $secondsSinceLastInputInfo seconds ago. Resetting idle counter. [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]" "INFO"
						Write-Host-Wrapper " " "INFO"
					}
				#}
				$hasSustainedActivity = $true
				$inputBasedActivityThisFrame = $true
			
				$script:g_ScreenSaverStarted = $false
				$script:g_DisplayTurnedOff = $false
				
				if(Test-IsSessionUnlocked -eq $true) {
					$script:g_PcLockedOnDemand = $false
				}
			}
			
			
			# Log to file sometimes, and to console some other times
			# -or ($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds'] -and $script:g_nextSettingsPollSeconds -le $script:Config['ActivityDetectionPeriodSamples']) 
			if (($script:g_idleSeconds -ge $script:Config['LogToFileIntervalSeconds'] -and $hasSustainedActivity -eq $true) -or $checkedSettings -eq $true) {
				$mouseLog = " "
				if (Test-IsInteractiveSession) {
					#$mouseLog = "MouseDelta: $([math]::Round($mouseDelta,1)) px"
					$mouseLog = "Input: $secondsSinceLastInputInfo s ago"
				} else {
					$mouseLog = "(mouse check skipped)"
				}
				$statusMessage = "CPU: $cpu % | GPU: $gpu % | Disk: $disk KBps | Net: $net KBps | $mouseLog | Idle: $([math]::Round($script:g_idleSeconds/60,3)) min | (T Sleep: $script:g_CurrentSleepIdleTimeMinutes min | T Hibernate: $script:g_CurrentHibernateIdleTimeMinutes min | T Display: $script:g_DisplayTimeoutDurationMinutes min | T ScreenSaver: $script:g_ScreensaverTimeoutDurationMinutes min | T Demand Win Lock: $($script:Config['LockPcAtThisIdleTimeMinutes']) min). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][idleSeconds_userOrAudioActivity: $([math]::Round($script:g_idleSeconds_userOrAudioActivity, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]"
				Write-Log $statusMessage "INFO"
			} elseif ($script:Config['LogToConsoleVerbose']) {
				$mouseLog = " "
				if (Test-IsInteractiveSession) {
					#$mouseLog = "MouseDelta: $([math]::Round($mouseDelta,1)) px"
					$mouseLog = "Input: $secondsSinceLastInputInfo s ago"
				} else {
					$mouseLog = "(mouse check skipped)"
				}
				$statusMessage = "CPU: $cpu % | GPU: $gpu % | Disk: $disk KBps | Net: $net KBps | $mouseLog | Idle: $([math]::Round($script:g_idleSeconds/60,3)) min | (T Sleep: $script:g_CurrentSleepIdleTimeMinutes min | T Hibernate: $script:g_CurrentHibernateIdleTimeMinutes min | T Display: $script:g_DisplayTimeoutDurationMinutes min | T ScreenSaver: $script:g_ScreensaverTimeoutDurationMinutes min | T Demand Win Lock: $($script:Config['LockPcAtThisIdleTimeMinutes']) min). [idleSeconds: $([math]::Round($script:g_idleSeconds, 5))][idleSeconds_userOrAudioActivity: $([math]::Round($script:g_idleSeconds_userOrAudioActivity, 5))][deltaTime: $([math]::Round($deltaTimeSeconds, 5))]"
				Write-Host-Wrapper $statusMessage "INFO"
			}

			if ($hasSustainedActivity) {
				$script:g_idleSeconds = 0.0
				if ($script:g_systemLoggedAsBeingIdle -eq $true) {
					LogSystemEvent_IdleOff
				}
				$script:g_cpuHistory.Clear()
				$script:g_diskHistory.Clear()
				$script:g_netHistory.Clear()
				$script:g_audioHistory.Clear()
			} else {
				$script:g_idleSeconds += $deltaTimeSeconds
			}
			
			$userActivity = $false
			if ($inputBasedActivityThisFrame -or $audioBasedActivityThisFrame) {
				$userActivity = $true
				$script:g_idleSeconds_userOrAudioActivity = 0.0
				if ($script:g_systemLoggedAsBeingIdle -eq $true) {
					LogSystemEvent_IdleOff
				}
			} else {
				$script:g_idleSeconds_userOrAudioActivity += $deltaTimeSeconds
			}
			
			# Hibernate or sleep if applicable under a valid battery user set threshold.
			$script:g_batteryLevel = GetBatteryLevel
			$script:g_isPluggedIn = IsComputerPluggedIn
			$forceHibernateTime = $script:Config['HibernateWhenBatteryIsBelowLevel']
			if($script:g_isPluggedIn -eq $true) {
				# NOTE: This is an important hack to rely on because laptops nowadays can have an alternative (slower) usb charger, which can cause the laptop battery to drain even if and while it's plugged in.
				$forceHibernateTime = $script:Config['HibernateWhenBatteryIsBelowLevel'] / 2.0
			}
			$forceSleepTime = $script:Config['SleepWhenBatteryIsBelowLevel']
			if($script:g_isPluggedIn -eq $true) {
				# NOTE: This is an important hack to rely on because laptops nowadays can have an alternative (slower) usb charger, which can cause the laptop battery to drain even if and while it's plugged in.
				$forceSleepTime = $script:Config['SleepWhenBatteryIsBelowLevel'] / 2.0
			}
			
			if ($script:Config['HibernateWhenBatteryIsBelowLevel'] -gt 0 -and $script:g_batteryLevel -lt $forceHibernateTime) {
				$blb = $forceHibernateTime
				$batteryMessage = "Battery: $script:g_batteryLevel is below the HibernateWhenBatteryIsBelowLevel of $blb!"
				$abort = Show-AbortDialog -Seconds $script:Config['CriticalBatteryAbortWindowCountdownSeconds'] -ActionName "hibernate" -Message $batteryMessage
				
				if ($abort) {
					Write-Log "$batteryMessage Wanted to hibernate but user aborted!" "INFO"
					
					$script:g_idleSeconds = 0.0
					$script:g_idleSeconds_userOrAudioActivity = 0.0
					if ($script:g_systemLoggedAsBeingIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
				} else {
					Write-Host-Wrapper "$batteryMessage Hibernating!" "INFO"
					Enter-HibernateState
					
					$script:g_idleSeconds = 0.0
					$script:g_idleSeconds_userOrAudioActivity = 0.0
					if ($script:g_systemLoggedAsBeingIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
				}
			} elseif ($script:Config['SleepWhenBatteryIsBelowLevel'] -gt 0 -and $script:g_batteryLevel -lt $forceSleepTime) {
				$blb = $forceSleepTime
				$batteryMessage = "Battery: $script:g_batteryLevel is below the SleepWhenBatteryIsBelowLevel of $blb!"
				$abort = Show-AbortDialog -Seconds $script:Config['CriticalBatteryAbortWindowCountdownSeconds'] -ActionName "sleep" -Message $batteryMessage
				
				if ($abort) {
					Write-Log "$batteryMessage Wanted to sleep but user aborted!" "INFO"
					
					$script:g_idleSeconds = 0.0
					$script:g_idleSeconds_userOrAudioActivity = 0.0
					if ($script:g_systemLoggedAsBeingIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
				} else {
					Write-Host-Wrapper "$batteryMessage Sleeping!" "INFO"
					Enter-SleepState
					
					$script:g_idleSeconds = 0.0
					$script:g_idleSeconds_userOrAudioActivity = 0.0
					if ($script:g_systemLoggedAsBeingIdle -eq $true) {
						LogSystemEvent_IdleOff
					}
				}
			}

			if($inputBasedActivityThisFrame -eq $false) {
				# Check if we are in charge of turning off the display or turning on any screensaver, and do it if it's time 
				$displayIdleMinutes = $script:g_idleSeconds / 60
				$iaflag = $script:Config['UseOnlyInputAndAudioEventsForDisplayOff']
				if ($userActivity -eq $false -and $iaflag -eq $true) {
					$displayIdleMinutes = $script:g_idleSeconds_userOrAudioActivity / 60
				}

				if ($script:Config['PreventAndReplaceWindowsAutoSleep']) {
					if($script:g_ScreenSaverStarted -eq $false -and $script:g_ScreensaverTimeoutDurationMinutes -and $script:g_ScreensaverTimeoutDurationMinutes -gt $script:Config['FailsafeTimeMinutes'] -and $displayIdleMinutes -gt $script:g_ScreensaverTimeoutDurationMinutes) {
						#Write-Log ">>>> starting screensaver (pwas): iaflag: $iaflag, displayIdleMinutes: $displayIdleMinutes, script:g_DisplayTimeoutDurationMinutes: $script:g_DisplayTimeoutDurationMinutes)" "INFO"
						Start-Screensaver
					}
					if($script:g_DisplayTurnedOff -eq $false -and $script:g_DisplayTimeoutDurationMinutes -and $script:g_DisplayTimeoutDurationMinutes -gt $script:Config['FailsafeTimeMinutes'] -and $displayIdleMinutes -gt $script:g_DisplayTimeoutDurationMinutes) {
						#Write-Log ">>>> turning off display (pwas): iaflag: $iaflag, displayIdleMinutes: $displayIdleMinutes, script:g_DisplayTimeoutDurationMinutes: $script:g_DisplayTimeoutDurationMinutes)" "INFO"
						Turn-Display-Off
					}
				} else {
					if($script:Config['TurnOnScreensaverAtThisIdleTimeMinutes'] -gt 0.0 -and $script:g_ScreenSaverStarted -eq $false -and $script:g_ScreensaverTimeoutDurationMinutes -and $script:g_ScreensaverTimeoutDurationMinutes -gt $script:Config['FailsafeTimeMinutes'] -and $displayIdleMinutes -gt $script:g_ScreensaverTimeoutDurationMinutes) {
						#Write-Log ">>>> starting screensaver: iaflag: $iaflag, displayIdleMinutes: $displayIdleMinutes, script:g_DisplayTimeoutDurationMinutes: $script:g_DisplayTimeoutDurationMinutes)" "INFO"
						Start-Screensaver
					}
					if($script:Config['TurnOffDisplayAtThisIdleTimeMinutes'] -gt 0.0 -and $script:g_DisplayTurnedOff -eq $false -and $script:g_DisplayTimeoutDurationMinutes -and $script:g_DisplayTimeoutDurationMinutes -gt $script:Config['FailsafeTimeMinutes'] -and $displayIdleMinutes -gt $script:g_DisplayTimeoutDurationMinutes) {
						#Write-Log ">>>> turning off display: iaflag: $iaflag, displayIdleMinutes: $displayIdleMinutes, script:g_DisplayTimeoutDurationMinutes: $script:g_DisplayTimeoutDurationMinutes)" "INFO"
						Turn-Display-Off
					}
				}
				# else NOTE: if $script:Config['FollowTheSameSleepAndScreenTimeSettingAsYourPowerPlan'] is true and $script:Config['PreventAndReplaceWindowsAutoSleep'] is false, then the PC will turn off its display and turn on its screensaver on its own, no involvement from us
				
				if ($script:g_PcLockedOnDemand -eq $false -and $script:Config['LockPcAtThisIdleTimeMinutes'] -and $script:Config['LockPcAtThisIdleTimeMinutes'] -gt $script:Config['FailsafeTimeMinutes'] -and $displayIdleMinutes -gt $script:Config['LockPcAtThisIdleTimeMinutes']) {
					# Write-Log "g_PcLockedOnDemand: $script:g_PcLockedOnDemand, LockPcAtThisIdleTimeMinutes: $script:Config['LockPcAtThisIdleTimeMinutes'], FailsafeTimeMinutes: $script:Config['FailsafeTimeMinutes'], idleSeconds: $script:g_idleSeconds" "INFO"
					Write-Log "Locking PC on demand at $($script:Config['LockPcAtThisIdleTimeMinutes']) min of idle." "Info"
					Lock-PC
				}
			}

			# Note: this doesn't work unless you run the script as administrator, so I commented it out ctrl+f:[respectOtherApps]
			#if ($RespectOtherAppsSleepExecutionPreventionFlags -eq $false -or (Test-OtherSystemExecutionStateHeld -eq $false -and $RespectOtherAppsSleepExecutionPreventionFlags -eq $true)) {
				# Check if ready to sleep
				if ($script:g_CurrentSleepIdleTimeMinutes -gt 0.0 -and $script:g_idleSeconds -ge ($script:g_CurrentSleepIdleTimeMinutes * 60)) {
					if($script:sleepOrHibernatePreventionFlagExists -eq $false) {
						$abort = Show-AbortDialog -Seconds $script:Config['AbortWindowCountdownSeconds'] -ActionName "sleep" -Message "Your PC has been idle for $script:g_CurrentSleepIdleTimeMinutes minutes."

						if ($abort) {
							Write-Log "User aborted sleep." "INFO"
							$script:g_idleSeconds = 0.0
							$script:g_idleSeconds_userOrAudioActivity = 0.0
							if ($script:g_systemLoggedAsBeingIdle -eq $true) {
								LogSystemEvent_IdleOff
							}
						} else {
							Write-Log "Proceeding to sleep because no abort or non-interactive session..." "INFO"
							#if (Test-Path-Timeouted -timeoutMilliseconds 6000 -pathToTest $script:Config['DontSleepWhileThisFileExistsPath']) {
							#    Remove-Item $script:Config['DontSleepWhileThisFileExistsPath'] -Force
							#    Write-Log "Deleted flag file: $script:Config['DontSleepWhileThisFileExistsPath']" "INFO"
							#}
							
							Enter-SleepState
							$script:g_idleSeconds = 0.0
							$script:g_idleSeconds_userOrAudioActivity = 0.0
							if ($script:g_systemLoggedAsBeingIdle -eq $true) {
								LogSystemEvent_IdleOff
							}
							Write-Log "System be woke. Resuming monitoring." "INFO"
						}
					}
				}
				# Check if ready to hibernate
				if ($script:g_CurrentHibernateIdleTimeMinutes -gt 0.0 -and $script:g_idleSeconds -ge ($script:g_CurrentHibernateIdleTimeMinutes * 60)) {
					if($script:sleepOrHibernatePreventionFlagExists -eq $false) {
						$abort = Show-AbortDialog -Seconds $script:Config['AbortWindowCountdownSeconds'] -ActionName "hibernate" -Message "Your PC has been idle for $script:g_CurrentSleepIdleTimeMinutes minutes."

						if ($abort) {
							Write-Log "User aborted hibernate." "INFO"
							$script:g_idleSeconds = 0.0
							$script:g_idleSeconds_userOrAudioActivity = 0.0
							if ($script:g_systemLoggedAsBeingIdle -eq $true) {
								LogSystemEvent_IdleOff
							}
						} else {
							Write-Log "Proceeding to hibernate because no abort or non-interactive session..." "INFO"
							#if (Test-Path-Timeouted -timeoutMilliseconds 6000 -pathToTest $script:Config['DontSleepWhileThisFileExistsPath']) {
							#    Remove-Item $script:Config['DontSleepWhileThisFileExistsPath'] -Force
							#    Write-Log "Deleted flag file: $script:Config['DontSleepWhileThisFileExistsPath']" "INFO"
							#}
							
							Enter-HibernateState
							$script:g_idleSeconds = 0.0
							$script:g_idleSeconds_userOrAudioActivity = 0.0
							if ($script:g_systemLoggedAsBeingIdle -eq $true) {
								LogSystemEvent_IdleOff
							}
							Write-Log "System be woke. Resuming monitoring." "INFO"
						}
					}
				}
			#}
		}
		
		# calculate a delta time again to measure how long it took to calculate this frame.
		$currentElapsedSeconds = $script:g_sw.Elapsed.TotalSeconds
		$deltaTimeSeconds = $currentElapsedSeconds - $script:g_lastElapsedSeconds
		
		
		# Using delta time throughout the loop we account for any lag spikes, and here if we finish faster we wait until the 1s mark. Also elsewhere we account (reset idle) for any long time spike (> 1 min) from an eventual system sleep and resume, or just system freeze.
		Start-Sleep -Milliseconds ([int]([math]::Max(0, 1000-$deltaTimeSeconds*1000)))
		
	}
} finally {
	if($script:Config['PreventAndReplaceWindowsAutoSleep'] -eq $true){
		$script:g_PreventSleep_ES = $false
		[WindowsSleepWrangler]::StopIgnoringIdleTimers()
	}
	if ($script:g_systemLoggedAsBeingIdle -eq $true) {
		LogSystemEvent_IdleOff
	}
	
	Write-Log "~~~~~~~~~ Exited the WICK script started at $script:StartTime PID: $PID ~~~~~~~*-" "INFO"
	Write-Log " " "INFO"
	LogSystemEvent_OnEnd
}