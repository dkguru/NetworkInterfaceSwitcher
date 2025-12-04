using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Windows.Forms;

namespace NetworkInterfaceSwitcher
{
    public partial class MainForm : Form
    {
        Version v;

        private ComboBox cmbInterface1;
        private ComboBox cmbInterface2;
        private Button btnSwitch;
        private Label lblStatus;
        private Label lblVersion;
        private Label lblInterface1;
        private Label lblInterface2;
        private Button btnRefresh;

        private Label lblInterface1Status;
        private Label lblInterface2Status;

        private NotifyIcon notifyIcon;  // Icon in the system tray
        private System.Windows.Forms.Timer statusTimer;

        // Registry Path for Storing Selection:  HKCU\Software\NetworkInterfaceSwitcher
        private const string RegistryRoot = @"Software\NetworkInterfaceSwitcher";

        public MainForm()
        {
            InitializeComponent();
            LoadNetworkInterfaces();
        }

        private void InitializeComponent()
        {
            this.Text = "Network Interface Switcher";
            this.Size = new System.Drawing.Size(500, 300);
            this.StartPosition = FormStartPosition.CenterScreen;

            lblInterface1 = new Label
            {
                Text = "Interface 1:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(100, 20)
            };

            cmbInterface1 = new ComboBox
            {
                Location = new System.Drawing.Point(130, 20),
                Size = new System.Drawing.Size(320, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbInterface1.SelectedIndexChanged += (s, e) => UpdateInterfaceStatus();

            lblInterface1Status = new Label
            {
                Location = new System.Drawing.Point(460, 20),
                Size = new System.Drawing.Size(20, 20),
                Text = "●",
                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.Gray
            };

            lblInterface2 = new Label
            {
                Text = "Interface 2:",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(100, 20)
            };

            cmbInterface2 = new ComboBox
            {
                Location = new System.Drawing.Point(130, 60),
                Size = new System.Drawing.Size(320, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbInterface2.SelectedIndexChanged += (s, e) => UpdateInterfaceStatus();

            lblInterface2Status = new Label
            {
                Location = new System.Drawing.Point(460, 60),
                Size = new System.Drawing.Size(20, 20),
                Text = "●",
                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.Gray
            };

            btnSwitch = new Button
            {
                Text = "Switch Interfaces",
                Location = new System.Drawing.Point(130, 110),
                Size = new System.Drawing.Size(150, 35)
            };
            btnSwitch.Click += BtnSwitch_Click;

            btnRefresh = new Button
            {
                Text = "Refresh List",
                Location = new System.Drawing.Point(300, 110),
                Size = new System.Drawing.Size(150, 35)
            };
            btnRefresh.Click += BtnRefresh_Click;

            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(20, 170),
                Size = new System.Drawing.Size(450, 60),
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            v = Assembly.GetExecutingAssembly().GetName().Version;
            // txtVersion.Text = $"v{v.Major}.{v.Minor}.{v.Build}.{v.Revision}";

            lblVersion = new Label
            {
                //Text = "(Version)",
                Text = $"v{v.Major}.{v.Minor}.{v.Build}.{v.Revision}",
                Location = new System.Drawing.Point(20, 235),
                Size = new System.Drawing.Size(450, 20),
                BorderStyle = BorderStyle.None,
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };

            this.Controls.Add(lblInterface1);
            this.Controls.Add(cmbInterface1);
            this.Controls.Add(lblInterface1Status);

            this.Controls.Add(lblInterface2);
            this.Controls.Add(cmbInterface2);
            this.Controls.Add(lblInterface2Status);

            this.Controls.Add(btnSwitch);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(lblStatus);
            this.Controls.Add(lblVersion);

            // Initialize NotifyIcon
            notifyIcon = new NotifyIcon
            {
                // Icon = new System.Drawing.Icon(GetType(), "NetworkInterfaceSwitcher"),
                Icon = SystemIcons.Application, // You can use a custom icon here
                Text = "Network Interface Switcher",
                Visible = false
            };
            notifyIcon.DoubleClick += NotifyIcon_DoubleClick;

            // Create context menu for tray icon
            ContextMenuStrip trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Open", null, TrayMenu_Open);
            trayMenu.Items.Add("Switch Interfaces", null, TrayMenu_Switch);
            trayMenu.Items.Add("-"); // Separator
            trayMenu.Items.Add("Exit", null, TrayMenu_Exit);
            notifyIcon.ContextMenuStrip = trayMenu;

            // Handle form resize to minimize to tray
            this.Resize += MainForm_Resize;

            // Update timer for status update
            statusTimer = new System.Windows.Forms.Timer();
            statusTimer.Interval = 5000; // 5 seconds
            statusTimer.Tick += (s, e) => UpdateInterfaceStatus();
            statusTimer.Start();
        }


        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon.Visible = true;
                notifyIcon.ShowBalloonTip(1000, "Network Interface Switcher",
                    "Application minimized to tray", ToolTipIcon.Info);
            }
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            ShowForm();
        }

        private void ShowForm()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();
            notifyIcon.Visible = false;
        }

        private void TrayMenu_Open(object sender, EventArgs e)
        {
            ShowForm();
        }

        private void TrayMenu_Switch(object sender, EventArgs e)
        {
            // Perform switch without showing the form
            BtnSwitch_Click(sender, e);
        }

        private void TrayMenu_Exit(object sender, EventArgs e)
        {
            notifyIcon.Visible = false;
            Application.Exit();
        }
        
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSettings();

            // Clean up the system refresh timer
            if (statusTimer != null)
            {
                statusTimer.Stop();
                statusTimer.Dispose();
            }

            // Clean up the tray icon
            if (notifyIcon != null)
            {
                notifyIcon.Visible = false;
                notifyIcon.Dispose();
            }

            base.OnFormClosing(e);
        }

        private void LoadSettings()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryRoot))
            {
                if (key == null) return;   // first run, nothing saved yet

                string saved1 = key.GetValue("Interface1") as string;
                string saved2 = key.GetValue("Interface2") as string;

                // Re-select the saved items if they still exist
                if (!string.IsNullOrEmpty(saved1))
                    cmbInterface1.SelectedItem = cmbInterface1.Items
                                                    .Cast<object>()
                                                    .FirstOrDefault(i => i.ToString() == saved1);

                if (!string.IsNullOrEmpty(saved2))
                    cmbInterface2.SelectedItem = cmbInterface2.Items
                                                    .Cast<object>()
                                                    .FirstOrDefault(i => i.ToString() == saved2);
            }
        }
        
        private void SaveSettings()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryRoot))
            {
                key.SetValue("Interface1", cmbInterface1.SelectedItem?.ToString() ?? string.Empty, RegistryValueKind.String);
                key.SetValue("Interface2", cmbInterface2.SelectedItem?.ToString() ?? string.Empty, RegistryValueKind.String);
            }
        }

        private void LoadNetworkInterfaces()
        {
            try
            {
                cmbInterface1.Items.Clear();
                cmbInterface2.Items.Clear();

                ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                    "SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID IS NOT NULL");

                foreach (ManagementObject adapter in searcher.Get())
                {
                    string name = adapter["NetConnectionID"]?.ToString();
                    if (!string.IsNullOrEmpty(name))
                    {
                        cmbInterface1.Items.Add(name);
                        cmbInterface2.Items.Add(name);
                    }
                }

                if (cmbInterface1.Items.Count > 0)
                    cmbInterface1.SelectedIndex = 0;
                if (cmbInterface2.Items.Count > 1)
                    cmbInterface2.SelectedIndex = 1;

                lblStatus.Text = $"Loaded {cmbInterface1.Items.Count} network interfaces";

                LoadSettings();
                UpdateInterfaceStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading interfaces: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateInterfaceStatus()
        {
            try
            {
                if (cmbInterface1.SelectedItem != null)
                {
                    string interface1 = cmbInterface1.SelectedItem.ToString();
                    bool interface1Enabled = IsInterfaceEnabled(interface1);

                    lblInterface1Status.ForeColor = interface1Enabled
                        ? System.Drawing.Color.Green
                        : System.Drawing.Color.Red;
                    lblInterface1Status.Text = interface1Enabled ? "●" : "●";

                    // Update tooltip
                    ToolTip tooltip1 = new ToolTip();
                    tooltip1.SetToolTip(lblInterface1Status,
                        interface1Enabled ? "Active" : "Disabled");
                }

                if (cmbInterface2.SelectedItem != null)
                {
                    string interface2 = cmbInterface2.SelectedItem.ToString();
                    bool interface2Enabled = IsInterfaceEnabled(interface2);

                    lblInterface2Status.ForeColor = interface2Enabled
                        ? System.Drawing.Color.Green
                        : System.Drawing.Color.Red;
                    lblInterface2Status.Text = interface2Enabled ? "●" : "●";

                    // Update tooltip
                    ToolTip tooltip2 = new ToolTip();
                    tooltip2.SetToolTip(lblInterface2Status,
                        interface2Enabled ? "Active" : "Disabled");
                }
            }
            catch (Exception ex)
            {
                // Silently handle errors in status update
                lblInterface1Status.ForeColor = System.Drawing.Color.Gray;
                lblInterface2Status.ForeColor = System.Drawing.Color.Gray;
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            LoadNetworkInterfaces();
            UpdateInterfaceStatus();
        }

        private void BtnSwitch_Click(object sender, EventArgs e)
        {
            if (cmbInterface1.SelectedItem == null || cmbInterface2.SelectedItem == null)
            {
                MessageBox.Show("Please select both interfaces", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbInterface1.SelectedItem.ToString() == cmbInterface2.SelectedItem.ToString())
            {
                MessageBox.Show("Please select different interfaces", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string interface1 = cmbInterface1.SelectedItem.ToString();
            string interface2 = cmbInterface2.SelectedItem.ToString();

            lblStatus.Text = "Switching interfaces...";
            Application.DoEvents();

            try
            {
                // Check current status
                bool interface1Enabled = IsInterfaceEnabled(interface1);
                bool interface2Enabled = IsInterfaceEnabled(interface2);

                // Perform the switch
                if (interface1Enabled)
                {
                    DisableInterface(interface1);
                    EnableInterface(interface2);
                    lblStatus.Text = $"Disabled: {interface1}\nEnabled: {interface2}";
                }
                else
                {
                    DisableInterface(interface2);
                    EnableInterface(interface1);
                    lblStatus.Text = $"Enabled: {interface1}\nDisabled: {interface2}";
                }

                SaveSettings();   // <-- record the last selections
                UpdateInterfaceStatus();
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"Error: {ex.Message}";
                MessageBox.Show($"Error switching interfaces: {ex.Message}\n\n" +
                    "Make sure you run this application as Administrator!",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool IsInterfaceEnabled(string interfaceName)
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                $"SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID = '{interfaceName}'");

            foreach (ManagementObject adapter in searcher.Get())
            {
                return (ushort)adapter["NetConnectionStatus"] == 2; // 2 = Connected
            }
            return false;
        }

        private void EnableInterface(string interfaceName)
        {
            ExecuteNetshCommand($"interface set interface \"{interfaceName}\" enable");
        }

        private void DisableInterface(string interfaceName)
        {
            ExecuteNetshCommand($"interface set interface \"{interfaceName}\" disable");
        }

        private void ExecuteNetshCommand(string arguments)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "netsh",
                Arguments = arguments,
                Verb = "runas", // Request admin privileges
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            using (Process process = Process.Start(psi))
            {
                process.WaitForExit();
                System.Threading.Thread.Sleep(1000); // Wait for interface to change state
            }
        }
    }

    //static class Program
    //{
    //    [STAThread]
    //    static void Main()
    //    {
    //        Application.EnableVisualStyles();
    //        Application.SetCompatibleTextRenderingDefault(false);
    //        Application.Run(new MainForm());
    //    }
    //}
}

namespace NetworkInterfaceSwitcher
{
    public partial class NetworkInterfaceSwitcher : Form
    {
        public NetworkInterfaceSwitcher()
        {
            InitializeComponent();
        }
    }
}
