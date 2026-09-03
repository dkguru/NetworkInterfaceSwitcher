namespace NetworkInterfaceSwitcher
{
    // Shared between the UI client and the service's named-pipe server so a "switch" request
    // can be handled entirely by the LocalSystem service, without the UI ever needing elevation.
    internal static class SwitchPipeContract
    {
        public const string PipeName = "NetworkInterfaceSwitcherPipe";
    }
}
