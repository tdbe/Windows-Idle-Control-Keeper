# If you get errors you might need to install:
# pip install pycaw
import sys
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IAudioMeterInformation

def is_audio_playing():
	"""Check if audio is actually playing by monitoring audio levels"""
	# Threshold tuned for Win10/11 idle noise floor. 
	# 0.01-0.03 filters out DC offset/micro-ticks without missing quiet playback.
	threshold = 0.02
	print(f"Checking for the first Peak > a threshold of {threshold}, among all the audio sessions.", file=sys.stderr) # printing to stderr to not affect output parsing from powershell
	# Get all active audio sessions (this would be all the bars in the classic windows volume mixer)
	sessions = AudioUtilities.GetAllSessions()

	for session in sessions:
		#if session.Process and session.Process.name() not in ['python.exe', 'pythonw.exe']:
		if session.Process:
			# Get audio meter information
			volume = session._ctl.QueryInterface(IAudioMeterInformation)
			peak = volume.GetPeakValue()
			# print(f"Checking audio for {session.Process}: Peak = {peak}")
			if peak > threshold:
				print(f"Detected audio playing in {session.Process}", file=sys.stderr) # printing to stderr to not affect output parsing from powershell
				return True
	print("No active audio detected.", file=sys.stderr) # printing to stderr to not affect output parsing from powershell
	return False

def get_master_volume():
    # Default render endpoint (speakers/headphones)
    device = AudioUtilities.GetSpeakers()
    volume = device.EndpointVolume.QueryInterface(IAudioEndpointVolume)
    return volume.GetMasterVolumeLevel()


if __name__ == "__main__":
	isPlaying = is_audio_playing()
	print(isPlaying)
