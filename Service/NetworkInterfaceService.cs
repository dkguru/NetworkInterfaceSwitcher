using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace NetworkInterfaceSwitcher.Service
{
    public class NetworkInterfaceService : ServiceBase
    {
        private System.Threading.Timer _timer;
        private CancellationTokenSource _pipeCts;
        private Task _pipeListenerTask;
        private readonly object _switchLock = new object();

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

            // Listens for immediate switch requests from the (unelevated) UI. This is what lets
            // the UI switch interfaces without ever prompting for UAC: the UI never touches netsh
            // itself, it just asks this LocalSystem-privileged service to do it.
            _pipeCts = new CancellationTokenSource();
            _pipeListenerTask = Task.Run(() => PipeListenLoopAsync(_pipeCts.Token));

            // Self-healing enforcement in case the adapter state drifts (e.g. someone toggles an
            // adapter from Windows' own UI) between pipe requests.
            _timer = new System.Threading.Timer(TimerCallback, null, TimeSpan.Zero, TimeSpan.FromSeconds(30));
        }

        protected override void OnStop()
        {
            Log("Service stopping");

            _timer?.Dispose();

            try
            {
                _pipeCts?.Cancel();
                _pipeListenerTask?.Wait(TimeSpan.FromSeconds(5));
            }
            catch { }
            finally
            {
                _pipeCts?.Dispose();
            }
        }

        private async Task PipeListenLoopAsync(CancellationToken token)
        {
            PipeSecurity pipeSecurity = BuildPipeSecurity();

            while (!token.IsCancellationRequested)
            {
                NamedPipeServerStream server = null;
                try
                {
                    server = NamedPipeServerStreamAcl.Create(
                        SwitchPipeContract.PipeName,
                        PipeDirection.InOut,
                        1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous,
                        4096,
                        4096,
                        pipeSecurity);

                    await server.WaitForConnectionAsync(token).ConfigureAwait(false);

                    using (var reader = new StreamReader(server, leaveOpen: true))
                    using (var writer = new StreamWriter(server, leaveOpen: true) { AutoFlush = true, NewLine = "\n" })
                    {
                        string response;
                        using (var requestCts = CancellationTokenSource.CreateLinkedTokenSource(token))
                        {
                            // Bound how long a single connected client can take to send its request,
                            // so a slow/stalled client can't tie up this single-instance pipe forever.
                            requestCts.CancelAfter(TimeSpan.FromSeconds(10));
                            try
                            {
                                // Protocol is 3 lines: command, interfaceA, interfaceB - one field per
                                // line, so an interface name can never be confused with a delimiter.
                                string command = await reader.ReadLineAsync(requestCts.Token).ConfigureAwait(false);
                                string interfaceA = await reader.ReadLineAsync(requestCts.Token).ConfigureAwait(false);
                                string interfaceB = await reader.ReadLineAsync(requestCts.Token).ConfigureAwait(false);
                                response = HandleRequest(command, interfaceA, interfaceB);
                            }
                            catch (OperationCanceledException) when (!token.IsCancellationRequested)
                            {
                                response = "ERROR|Timed out waiting for request";
                            }
                        }

                        await writer.WriteLineAsync(response).ConfigureAwait(false);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Log("Pipe listener error: " + ex);
                }
                finally
                {
                    try { server?.Disconnect(); } catch { }
                    server?.Dispose();
                }
            }
        }

        private static PipeSecurity BuildPipeSecurity()
        {
            var security = new PipeSecurity();

            // Only interactively logged-on sessions (console or Remote Desktop) can request a switch -
            // network logons and non-interactive service accounts are excluded. Only this service
            // (LocalSystem) actually performs the privileged netsh call.
            var interactiveUsers = new SecurityIdentifier(WellKnownSidType.InteractiveSid, null);
            security.AddAccessRule(new PipeAccessRule(interactiveUsers, PipeAccessRights.ReadWrite, AccessControlType.Allow));

            var localSystem = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
            security.AddAccessRule(new PipeAccessRule(localSystem, PipeAccessRights.FullControl, AccessControlType.Allow));

            var admins = new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null);
            security.AddAccessRule(new PipeAccessRule(admins, PipeAccessRights.FullControl, AccessControlType.Allow));

            return security;
        }

        private string HandleRequest(string command, string interfaceA, string interfaceB)
        {
            if (string.IsNullOrWhiteSpace(command))
                return "ERROR|Empty request";

            if (!string.Equals(command, "SWITCH", StringComparison.Ordinal))
                return "ERROR|Unknown command";

            if (string.IsNullOrWhiteSpace(interfaceA) || string.IsNullOrWhiteSpace(interfaceB) ||
                string.Equals(interfaceA, interfaceB, StringComparison.OrdinalIgnoreCase))
            {
                return "ERROR|Invalid interface names";
            }

            var result = HandleSwitchRequest(interfaceA, interfaceB);
            return (result.Success ? "OK|" : "ERROR|") + result.Message;
        }

        private (bool Success, string Message) HandleSwitchRequest(string interfaceA, string interfaceB)
        {
            lock (_switchLock)
            {
                bool aEnabled = IsInterfaceEnabled(interfaceA);
                string toEnable = aEnabled ? interfaceB : interfaceA;
                string toDisable = aEnabled ? interfaceA : interfaceB;

                var result = SwitchTo(toEnable, toDisable);
                if (result.Success)
                {
                    PersistState(interfaceA, interfaceB, toEnable);
                }
                return result;
            }
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
                    string active = key.GetValue("ActiveInterface") as string;

                    if (string.IsNullOrEmpty(iface1) || string.IsNullOrEmpty(iface2)) return;

                    if (string.Equals(iface1, iface2, StringComparison.OrdinalIgnoreCase))
                    {
                        Log($"Configured pair is invalid (Interface1 == Interface2 == '{iface1}'); skipping enforcement.");
                        return;
                    }

                    // Legacy fallback for installs that only ever set Interface1/Interface2 (e.g. via
                    // install-service.ps1) and never went through a pipe-driven switch yet.
                    if (string.IsNullOrEmpty(active) ||
                        (!string.Equals(active, iface1, StringComparison.OrdinalIgnoreCase) &&
                         !string.Equals(active, iface2, StringComparison.OrdinalIgnoreCase)))
                    {
                        active = iface1;
                    }

                    string other = string.Equals(active, iface1, StringComparison.OrdinalIgnoreCase) ? iface2 : iface1;

                    bool activeEnabled = IsInterfaceEnabled(active);
                    bool otherEnabled = IsInterfaceEnabled(other);

                    Log($"Configured pair: '{iface1}' / '{iface2}' - desired active: '{active}' - current states: active={activeEnabled} other={otherEnabled}");

                    if (activeEnabled && !otherEnabled)
                    {
                        return; // desired state already present
                    }

                    lock (_switchLock)
                    {
                        var result = SwitchTo(active, other);
                        Log($"Enforcement switch result: {result.Success} - {result.Message}");
                    }
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

        private (bool Success, string Message) SwitchTo(string interfaceToEnable, string interfaceToDisable)
        {
            // Enable the target before disabling the source, to minimize the window where both
            // interfaces are down (important if this service is itself being reached over one of them).
            var enableResult = ExecuteNetshCommand($"interface set interface \"{interfaceToEnable}\" enable");
            Log($"netsh enable '{interfaceToEnable}': exit={enableResult.ExitCode} stdout={enableResult.StdOut} stderr={enableResult.StdErr}");

            if (enableResult.ExitCode != 0)
            {
                // Don't touch interfaceToDisable if the target never came up - leaving the
                // original interface enabled is safer than ending up with neither enabled.
                return (false, $"netsh failed to enable '{interfaceToEnable}' (exit={enableResult.ExitCode}); left '{interfaceToDisable}' untouched");
            }

            var disableResult = ExecuteNetshCommand($"interface set interface \"{interfaceToDisable}\" disable");
            Log($"netsh disable '{interfaceToDisable}': exit={disableResult.ExitCode} stdout={disableResult.StdOut} stderr={disableResult.StdErr}");

            bool success = disableResult.ExitCode == 0;
            string message = success
                ? $"Enabled: {interfaceToEnable}; Disabled: {interfaceToDisable}"
                : $"Enabled '{interfaceToEnable}' but failed to disable '{interfaceToDisable}' (exit={disableResult.ExitCode})";

            return (success, message);
        }

        private void PersistState(string interface1, string interface2, string activeInterface)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey(RegistryRoot))
                {
                    key.SetValue("Interface1", interface1, RegistryValueKind.String);
                    key.SetValue("Interface2", interface2, RegistryValueKind.String);
                    key.SetValue("ActiveInterface", activeInterface, RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                Log("PersistState failed: " + ex);
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
        private const int NetshTimeoutMilliseconds = 15000;

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

                var stdOut = new StringBuilder();
                var stdErr = new StringBuilder();

                using (Process process = new Process { StartInfo = psi })
                {
                    // Read output asynchronously via events (not ReadToEnd) so a hung/silent netsh
                    // process can't block this thread before we ever get to the WaitForExit timeout.
                    process.OutputDataReceived += (s, e) => { if (e.Data != null) stdOut.AppendLine(e.Data); };
                    process.ErrorDataReceived += (s, e) => { if (e.Data != null) stdErr.AppendLine(e.Data); };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    if (!process.WaitForExit(NetshTimeoutMilliseconds))
                    {
                        Log($"netsh timed out after {NetshTimeoutMilliseconds}ms, killing process: {arguments}");
                        try { process.Kill(entireProcessTree: true); } catch (Exception killEx) { Log("Failed to kill timed-out netsh process: " + killEx); }
                        return (-1, stdOut.ToString(), $"netsh timed out after {NetshTimeoutMilliseconds}ms");
                    }

                    return (process.ExitCode, stdOut.ToString(), stdErr.ToString());
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
