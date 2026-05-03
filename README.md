# Windows-Idle-Control-Keeper

# Intro:

[Saul Goodman voice] **Can't get reliable sleep? Feeling like it's out of your control? Well fret not, just run this script and you can start counting those sheep!**

**WICK - Windows Idle Control Keeper**

This script detects Idle activity with your specific thresholds and conditions, and triggers / prevents Windows sleep, hibernate, display, lock, screen saver, on Your terms. Detects activity on CPU, (non-virtual) Network (both internet and LAN), Storage, mouse, and peak Sound value; counting idle-breaking event frequency and amplitude per time periods, to determine if an Idle timer should continue or be broken. It's Windows Power Plan aware, including display off and screensaver schedule, and maintains windows screen locking. No admin required.

I don't usually post my system scripts but it annoyed me that for such a wide need, there was nothing out there but forum threads of people using obscure and partial tools like [DontSleep!.exe](https://www.softwareok.com/?Download=DontSleep) [from 2014](https://www.chip.de/downloads/Don-t-Sleep_42626965.html)

# Features, Dependencies, Log Example, Notes, Run & Parameters:

## Features:

- Does not require administrator permission.
- Works even if windows is locked. Also works if logged out or never logged in (if you start it at system start via task scheduler).
- Shows warning / abort window for $AbortWindowCountdownSeconds before triggering a sleep or hibernate (if in an interactive session (not locked or logged out)).
- Dynamically reads (every minute) from your currently active windows power plan (plugged in or battery) to check sleep and hibernate times.
- Determines idle by accumulating sustained activity events over certain timeframes, through checks every second: based on if there's CPU, Network, sorage (without waking sleeping hard disks), audio, and mouse activity.
- Can prevent windows from sleeping until this script decides it's time to sleep.
- Can set a sleep or hibernate time for longer than 5h (the max that Windows power plan allows for some gormless reason).
- Allows a blacklist for logical drives e.g. `"L", "A", "N"` - you may have drives that have activity you consider passive and you're okay sleeping on. But also keep in mind the NetworkThresholdKBps setting.
- Can be paused while running by creating a `.ignore_running_Windows_Idle_Control_Keeper_script` flag file (e.g. renamed empty txt file).
- Logs what's going on to Windows' Event Viewer - Applicaton Log (only idle on (not immediate) and idle off), and logs to file at LogPath, so you know at what time Idle state was broken and after how much idle time. (or if there were errors) (log cleans itself up to stay less than LogMaxSizeMB)
- It maintains windows screen locking (also can lock on demand), and display off and screensaver schedule (can also trigger them on demand).

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
[2026-04-30 00:53:59] [INFO] CPU: 1 % | Disk: 8279 KBps | Net: 1201 KBps | Input: 91 s ago | Idle: 0 min | Sleep: 30 min | Hibernate: 60 min | Display Off: 0 sec | ScreenSaver: 0 sec
```

## Notes: 

### Note: 

Tested on Windows 11 LTSC, laptop and PC.

### Note:

This script cannot and will never: ask for administrator privileges, listen to your sounds, key presses, taps and clicks, connect to the internet or network. It only asks for, reads, and logs, when an event of a certain category happened, does not know what data it had.

### Note:

I built in a 60 second failsafe, before which this script won't do anything. So you can't screw yourself over e.g. set a 1s sleep timeout, or force lock etc. So you have 1 minute to fix it after wake or even on startup even if you set it to start at system start via task scheduler.

### Note: 

Because it doesn't use admin rights, while $PreventAndReplaceWindowsAutoSleep is $true, this script can't check if other processes requested that the system not sleep, e.g. an active remote desktop connection while the PC is otherwise within your idle thresholds; so it won't respect their request. You could fix this by either setting $PreventAndReplaceWindowsAutoSleep to $false, or by running this script as admin and uncomment the `ctrl+f:[respectOtherApps]` code blocks.

### Note:

I've always nuked Modern Standby from every PC I touched, because we have literally 0.0f low-power hardware and protocol standards, and I don't want constant 100W power draw, and for laptops my battery to run out in 2 hours while "sleeping" with the lid closed (Microsoft is the most infuriating thing in the history of ever). You can have a look at how I printed and fetched the sleep AC/DC settings, and figure out the parsing of anything else if you want. PRs welcome.

### Note: 

I wouldn't be caught dead writing for free 1000+ lines of powershell script of all things, so for this I tried out LLMs. I don't consider llm output even remotely reliable, but this is all verified and very re-written by me. For those curious: I used 256k context and: qwen 3 coder next 80b a3b q6, qwen 3.6 35b a3b q8, and qwen 3.6 27b q4, locally. They're "great" (within 5-10% of the scores of the huge frontier models) but simultaneously also completely shit at even such a simple job, and not just because this solution doesn't already exist: ie they picked network and storage checks that take at least 1s to return a value, and were calling them repeatedly in loops per disk and per adapter, resulting in a while loop that runs once every 7-10s.. So the verdict is I had to do all the thinking myself. They only oneshotted the logging, the cpu, the sleep functions, and the .PARAM list. Also the peak audio checking I had to research and write myself in python after many wildly off LLM solutions.

## Run & Parameters:

Parameters you can set when calling the script or adding it to Task Scheduler (the ones you skip will have defaults).

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