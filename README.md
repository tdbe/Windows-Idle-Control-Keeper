# Windows-Idle-Control-Keeper

# Intro:

[Saul Goodman voice] **Can't get reliable sleep? Feeling like it's out of your control? Well fret not, just run this script and you can start counting those sheep!**

**WICK - Windows Idle Control Keeper**

This script detects Idle activity with your specific thresholds and conditions, and triggers / prevents Windows Sleep on Your terms. Detects activity on CPU, (non-virtual) Network (both internet and LAN), Storage, mouse, and peak Sound value; using instances per time period, to determine if an Idle timer should continue or be broken. It does not affect and is not affected by (auto) windows screen locking, or (auto) turning off the display.

I don't usually post my system scripts but it annoyed me that for such a wide need, there was nothing out there but forum threads of people using ancient and partial tools like [DontSleep!.exe](https://www.softwareok.com/?Download=DontSleep) [from 2014](https://www.chip.de/downloads/Don-t-Sleep_42626965.html)

# Features, Dependencies, Log Example, Notes, Run & Parameters:

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
[2026-04-30 00:11:58] [INFO] ~*------- W.I.C.K. started. -------
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

### Note: 

I've always nuked Modern Standby from every PC I touched, because we have literally 0.0f low-power hardware and protocol standards, and I don't want constant 100W power draw, and for laptops my battery to run out in 2 hours while "sleeping" with the lid closed (Microsoft is the most infuriating thing in the history of ever). You can have a look at how I printed and fetched the sleep AC/DC settings, and figure out the parsing of anything else if you want. PRs welcome.

### Note: 

I don't vibecode anything I consider even remotely reliable, but here I tried out LLMs. (Don't worry it's all read-through, tested, and very re-written by me.) Otherwise I wouldn't be caught dead writing, for free, 700 lines of powershell script of all things. I used: qwen 3 coder next 80b a3b q6, qwen 3.6 35b a3b q8, and qwen 3.6 27b q4. They're "great" (within 5-10% of the scores of the huge frontier models) but simultaneously also completely shit at even such a simple job, and not just because this solution doesn't already exist: ie they picked network and storage checks that take at least 1s to return a value, and were calling them repeatedly in loops per disk and per adapter, resulting in a while loop that runs once every 7-10s.. So the verdict is I had to do all the thinking myself. They only oneshotted the logging, the cpu, the sleep functions, and the .PARAM list. Also the audio checking I had to research and write myself in python after many wildly off LLM solutions.

## Run & Parameters:

Parameters you can set when calling the script or adding it to Task Scheduler (the ones you skip will have defaults).

To see all parameters and their description, run this command: 

```
get-help "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -detailed
```

### Run in a powershell terminal window, examples:

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepTimeSettingAsYourPowerPlan:$true -FallbackIdleMinutes:30 -OnlyThisScriptCanCauseWindowsToSleep:$true # other flags -etc. -etc.

or

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepTimeSettingAsYourPowerPlan:$false -OnlyThisScriptCanCauseWindowsToSleep:$true -IdleDurationMinutes:720 -LogToConsoleVerbose:$false # other flags -etc. -etc.

```

### Path parameters to check and/or add modified versions:


```
.PARAMETER -PycawAudioCheckerPath
  Full path to the Python script used to detect audio playback.  
  (default: "C:\Commands_And_Logs\Pycaw_check_if_audio_is_playing.py")

.PARAMETER -PythonPath
  Full path to the Python executable used to run the audio checker script.
  (default: "$env:USERPROFILE\AppData\Local\Programs\Python\Python312\python.exe")

.PARAMETER -PauseFlagPath
  Path to a flag file. If this file exists, the script pauses monitoring and skips sleep.  
  Allows manual pause/resume by creating/deleting the file.
  (default: "C:\Command_And_Logs\.ignore_running_Windows_Idle_Control_Keeper_script")

.PARAMETER -LogPath
  Full path to the log file (default: "C:\Commands_And_Logs\Windows_Idle_Control_Keeper.log")
```

### Run in Task Scheduler:

#### Program/script: 

```
powershell.exe
```

#### Add arguments (window opens as minimized):

```
-NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepTimeSettingAsYourPowerPlan:$true -FallbackIdleMinutes:30 -OnlyThisScriptCanCauseWindowsToSleep:$true #  other flags -etc. -etc.
```

#### Add arguments (no window, runs in background completely hidden):

```
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:/Commands_And_Logs/windows_idle_control_keeper.ps1" -FollowTheSameSleepTimeSettingAsYourPowerPlan:$true -FallbackIdleMinutes:30 -OnlyThisScriptCanCauseWindowsToSleep:$true #  other flags -etc. -etc.
```
