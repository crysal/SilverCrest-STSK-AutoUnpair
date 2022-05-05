# SilverCrest-STSK-AutoUnpair
A powershell script with windows scheduler to unpair the bluetooth connection when power is lost


The silver crest bluetooth headphones using the STSK A1 model does not reconnect if power is lost and then regained, this requires windows users to go into bluetooth settins and unpair the device, just to pair it again.

This script and scheduler will detect when power is lost to the bluetooth headphones by looking at the status of the audio plug and play device, then send a unpair request to the bluetooth driver for any device with the name \*STSK*

It uses the *Microsoft-Windows-Audio/Operational, Microsoft-Windows-Audio, Event ID: 65* on the scheduler to trigger the powershell script

If you have the STSK A1 headset/headphones as the default playback device, then you need to change the audio everytime you connect the device, but this script will also test if you are connected and then set the volume to 50% (you can change this to what you want here https://github.com/crysal/SilverCrest-STSK-AutoUnpair/blob/b1d6a3adafed17e9d7cd3d23dfcc0f5e5e20ea3c/BlueToothUnpair.ps1#L70)
