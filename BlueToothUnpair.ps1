if ( (Get-PnpDevice -FriendlyName "*STSK*").status -eq "Unknown" ){
Write-Progress -Activity "UnPairing" -Status "Reading Dll Files" -PercentComplete 0 -CurrentOperation "Importing"
#Start-Process "ms-settings:connecteddevices"for 
$BTlib = @"
   [DllImport("BluetoothAPIs.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
   [return: MarshalAs(UnmanagedType.U4)]
   static extern UInt32 BluetoothRemoveDevice(IntPtr pAddress);
   public static UInt32 Unpair(UInt64 BTAddress) {
      GCHandle pinnedAddr = GCHandle.Alloc(BTAddress, GCHandleType.Pinned);
      IntPtr pAddress     = pinnedAddr.AddrOfPinnedObject();
      UInt32 result       = BluetoothRemoveDevice(pAddress);
      pinnedAddr.Free();
      return result;
   }
"@
Write-Progress -Activity "UnPairing" -Status "Searching Hardware" -PercentComplete 25 -CurrentOperation "Setting HardwareID"
$BTDevices = @(Get-PnpDevice -class bluetooth -FriendlyName "*STSK*" | select HardwareID, @{N='Address';E={[uInt64]('0x{0}' -f $_.HardwareID[0].Substring(12))}} | select Address)
Write-Progress -Activity "UnPairing" -Status "Creating Members" -PercentComplete 50 -CurrentOperation "BTRemover Created"
$BTR = Add-Type -MemberDefinition $BTlib -Name "BTRemover"  -Namespace "BStuff" -PassThru
Write-Progress -Activity "UnPairing" -Status "Trying to UnPair" -PercentComplete 75 -CurrentOperation "Sending unpair request"
$BTR::Unpair($BTDevices[0].Address)
Write-Progress -Activity "UnPairing" -Status "Unpaired" -PercentComplete 100 -CurrentOperation "Exiting"
exit 0
} elseif ((Get-PnpDevice -class AudioEndpoint -FriendlyName "*STSK*").status -eq "Ok" ) {
Write-Progress -Activity "Enabling Volume" -Status "Defining type" -PercentComplete 10 -CurrentOperation "Creating Interface"
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume
{
    // f(), g(), ... are unused COM method slots. Define these if you care
    int f(); int g(); int h(); int i();
    int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
    int j();
    int GetMasterVolumeLevelScalar(out float pfLevel);
    int k(); int l(); int m(); int n();
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice
{
    int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator
{
    int f(); // Unused
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio
{
    static IAudioEndpointVolume Vol()
    {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
        IAudioEndpointVolume epv = null;
        var epvid = typeof(IAudioEndpointVolume).GUID;
        Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
        return epv;
    }
    public static float Volume
    {
        get { float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v; }
        set { Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty)); }
    }
}
'@
Write-Progress -Activity "Enabling Volume" -Status "Changing Volume" -PercentComplete 75 -CurrentOperation "Setting volume to 50"
[audio]::Volume = 0.5
Write-Progress -Activity "Enabling Volume" -Status "Volume Enabled" -PercentComplete 100 -CurrentOperation "Exiting"
exit 0
}
exit 0
