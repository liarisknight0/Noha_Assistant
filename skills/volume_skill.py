from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

class VolumeSkill:
    def __init__(self):
        # These are the trigger words Noha will look for
        self.keywords = ["volume", "sound"]

    def can_handle(self, command):
        """Checks if the user's command is asking to change the volume."""
        return any(keyword in command.lower() for keyword in self.keywords)

    def execute(self, command, value=None):
        """Executes the volume change."""
        if value is not None:
            try:
                # Get the Windows Audio endpoints
                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))
                
                # Convert percentage (0-100) to scalar (0.0 to 1.0)
                try:
                    target_value = int(value)
                except ValueError:
                    print(f"Noha: Invalid volume value '{value}'. Please provide a number.")
                    return False
                
                target_value = min(max(target_value, 0), 100)
                target_vol = target_value / 100.0
                
                # Set the volume
                volume.SetMasterVolumeLevelScalar(target_vol, None)
                print(f"🔊 Noha: I have set the volume to {target_value}%.")
                return True
            except Exception as e:
                print(f"Noha: I ran into an error setting the volume: {e}")
                return False
        else:
            print("Noha: I heard you ask for volume, but I didn't catch the percentage number.")
            return False
