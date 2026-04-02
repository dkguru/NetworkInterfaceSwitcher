using System;
using System.Diagnostics;
using System.ServiceProcess;
using System.Threading;
using System.IO;
using Microsoft.Win32;

namespace NetworkInterfaceSwitcher.Service
{
    public class NetworkInterfaceService : ServiceBase
    {
        private System.Threading.Timer _timer;
        private const string RegistryRoot = @"SOFTWARE\NetworkInterfaceSwitcher";
        private readonly string _logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "NetworkInterfaceSwitcher");
        private readonly string _logFile;
        private readonly object _logLock = new object();

        public NetworkInterfaceService()
        {
            this.ServiceName = "NetworkInterfaceSwitcherService";
            _logFile = Path.Combine(_logDir, "service.log");
        }

        protected override void OnStart(string[] args)
        {
            // Ensure log directory
            try
            {
                Directory.CreateDirectory(_logDir);
            }
            catch { }

            Log("Service starting");

            // Run every 30 seconds
            _timer = new System.Threading.Timer(TimerCallback, null, TimeSpan.Zero, TimeSpan.FromSeconds(30));
        }

        protected override void OnStop()
        {
            Log("Service stopping");
            _timer?.Dispose();
        }

        private void TimerCallback(object state)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(RegistryRoot))
                {
                    if (key == null) return;

                    string iface1 = key.GetValue("Interface1") as string;
                    string iface2 = key.GetValue("Interface2") as string;

                    if (string.IsNullOrEmpty(iface1) || string.IsNullOrEmpty(iface2)) return;

                    // Determine current status and enforce the saved pair
                    bool iface1Enabled = IsInterfaceEnabled(iface1);
                    bool iface2Enabled = IsInterfaceEnabled(iface2);

                    Log($"Configured pair: '{iface1}' / '{iface2}' - current states: {iface1Enabled}/{iface2Enabled}");

                    if (iface1Enabled && !iface2Enabled)
                    {
                        Log("Desired state already present");
                        return; // desired state
                    }

                    // Try to enable iface1 and disable iface2
                    var out1 = ExecuteNetshCommand($"interface set interface \"{iface1}\" enable");
                    var out2 = ExecuteNetshCommand($"interface set interface \"{iface2}\" disable");
                    Log($"netsh enable '{iface1}': {out1.ExitCode} stdout={out1.StdOut} stderr={out1.StdErr}");
                    Log($"netsh disable '{iface2}': {out2.ExitCode} stdout={out2.StdOut} stderr={out2.StdErr}");
                }
            }
            catch (Exception ex)
            {
                // Log to event log to aid debugging
                try
                {
                    EventLog.WriteEntry("NetworkInterfaceSwitcherService", ex.ToString(), EventLogEntryType.Error);
                    Log("Exception in TimerCallback: " + ex);
                }
                catch { }
            }
        }

        // Helper methods copied from UI class but non-UI
        private bool IsInterfaceEnabled(string interfaceName)
        {
            try
            {
                var searcher = new System.Management.ManagementObjectSearcher(
                    $"SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID = '{interfaceName}'");

                foreach (System.Management.ManagementObject adapter in searcher.Get())
                {
                    return (ushort)adapter["NetConnectionStatus"] == 2; // 2 = Connected
                }
            }
            catch { }
            return false;
        }
        private (int ExitCode, string StdOut, string StdErr) ExecuteNetshCommand(string arguments)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "netsh",
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    string outStr = process.StandardOutput.ReadToEnd();
                    string errStr = process.StandardError.ReadToEnd();
                    process.WaitForExit();
                    return (process.ExitCode, outStr, errStr);
                }
            }
            catch (Exception ex)
            {
                Log("ExecuteNetshCommand exception: " + ex);
                return (-1, string.Empty, ex.ToString());
            }
        }

        private void Log(string message)
        {
            try
            {
                lock (_logLock)
                {
                    string text = $"[{DateTime.UtcNow:O}] {message}" + Environment.NewLine;
                    File.AppendAllText(_logFile, text);
                }
            }
            catch { }
        }
    }
}
