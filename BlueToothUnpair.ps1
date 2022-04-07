if ( (Get-PnpDevice -FriendlyName "*STSK*").status -eq "Unknown" ){
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
$BTDevices = @(Get-PnpDevice -class bluetooth -FriendlyName "*STSK*" | select HardwareID, @{N='Address';E={[uInt64]('0x{0}' -f $_.HardwareID[0].Substring(12))}} | select Address)
$BTR = Add-Type -MemberDefinition $BTlib -Name "BTRemover"  -Namespace "BStuff" -PassThru
$BTR::Unpair($BTDevices[0].Address)
}
