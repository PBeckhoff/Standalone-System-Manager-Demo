using System.Windows.Forms;
using System.Xml.Linq;
using TcSysManRMLib;

namespace UI_Disable_App
{
    public partial class MainForm : Form
    {
        private ITcSysManager15 sysManager;
        private TextBox txtOutput;
        private Button btnLoadConfig;
        private Button btnSaveConfig;
        private Button btnLoadXTI;
        private Button btnScanNetwork;
        private Button btnClearLog;
        private GroupBox grpActions;
        private FlowLayoutPanel flowDevices;
        private Label lblDevices;
        private Panel pnlDevices;
        private bool configurationLoaded = false;
        private string filePath = "";
        private System.Collections.Generic.Dictionary<CheckBox, ITcSmTreeItem> deviceCheckBoxes = new System.Collections.Generic.Dictionary<CheckBox, ITcSmTreeItem>();
        
        public MainForm()
        {
            InitializeComponent();
            InitializeTwinCAT();
        }

        private void InitializeComponent()
        {
            this.Text = "TwinCAT System Manager";
            this.Size = new System.Drawing.Size(1000, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Main panel to hold everything
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill
            };

            // Output TextBox
            txtOutput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9F)
            };
            // Devices Panel (Right side)
            pnlDevices = new Panel
            {
                Dock = DockStyle.Right,
                Width = 300,
                BorderStyle = BorderStyle.FixedSingle
            };
            lblDevices = new Label
            {
                Text = "Devices",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.LightGray
            };
            flowDevices = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(10)
            };
            pnlDevices.Controls.Add(flowDevices);
            pnlDevices.Controls.Add(lblDevices);
            // GroupBox for actions
            grpActions = new GroupBox
            {
                Text = "Actions",
                Dock = DockStyle.Top,
                Height = 80
            };
            // Load Configuration Button
            btnLoadConfig = new Button
            {
                Text = "Load Configuration",
                Location = new System.Drawing.Point(10, 25),
                Size = new System.Drawing.Size(120, 40)
            };
            btnLoadConfig.Click += BtnLoadConfig_Click;

            // Load XTI Button
            btnLoadXTI = new Button
            {
                Text = "Load XTI File",
                Location = new System.Drawing.Point(140, 25),
                Size = new System.Drawing.Size(120, 40),
                Enabled = false           
            };
            btnLoadXTI.Click += BtnLoadXTI_Click;

            // Save XTI Button
            btnSaveConfig = new Button
            {
                Text = "Save Configuration",
                Location = new System.Drawing.Point(530, 25),
                Size = new System.Drawing.Size(120, 40)
            };
            btnSaveConfig.Click += BtnSaveConfig_Click;

            // Scan Network Button
            btnScanNetwork = new Button
            {
                Text = "Scan Network",
                Location = new System.Drawing.Point(270, 25),
                Size = new System.Drawing.Size(120, 40),
                Enabled = false
            };
            btnScanNetwork.Click += BtnScanNetwork_Click;

            // Clear Log Button
            btnClearLog = new Button
            {
                Text = "Clear Log",
                Location = new System.Drawing.Point(400, 25),
                Size = new System.Drawing.Size(120, 40)
            };
            btnClearLog.Click += (s, e) => txtOutput.Clear();

            grpActions.Controls.Add(btnLoadConfig);
            //grpActions.Controls.Add(btnLoadXTI);
            //grpActions.Controls.Add(btnScanNetwork);
            grpActions.Controls.Add(btnClearLog);
            grpActions.Controls.Add(btnSaveConfig);

            mainPanel.Controls.Add(txtOutput);
            mainPanel.Controls.Add(pnlDevices);
            mainPanel.Controls.Add(grpActions);

            this.Controls.Add(mainPanel);
            this.FormClosing += MainForm_FormClosing;
        }

        private void InitializeTwinCAT()
        {
            try
            {
                var systemManagerRM = Activator.CreateInstance(Type.GetTypeFromProgID("TcSysManagerRM", throwOnError: true)!) as ITcSysManagerRM;
                sysManager = systemManagerRM?.CreateSysManager15();
                Log("TwinCAT System Manager initialized successfully.");
            }
            catch (Exception ex)
            {
                Log($"Error initializing TwinCAT: {ex.Message}");
                MessageBox.Show($"Failed to initialize TwinCAT System Manager.\n\n{ex.Message}",
                    "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnLoadConfig_Click(object sender, EventArgs e)
        {
            LoadConfiguration();
        }
        private void BtnSaveConfig_Click(object sender, EventArgs e)
        {
            SaveConfiguration();
        }
        private void BtnLoadXTI_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "XTI Files (*.xti)|*.xti|All Files (*.*)|*.*";
                openFileDialog.Title = "Select XTI Configuration File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    LoadXTIFile(openFileDialog.FileName);
                }
            }
        }
        private void SaveConfiguration()
        {
            sysManager.SaveConfiguration("");
            Log($"Configuration Saved");
        }
        private void LoadConfiguration()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "TwinCAT Project Files (*.tsproj)|*.tsproj|All Files (*.*)|*.*";
                openFileDialog.Title = "Select TwinCAT Project File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Log($"Loading TwinCAT project: {openFileDialog.FileName}");

                        // Open the TwinCAT project
                        sysManager.OpenConfiguration(openFileDialog.FileName);

                        Log("Configuration loaded successfully.");
                        Log("You can now load XTI files or scan for devices.");

                        configurationLoaded = true;
                        UpdateButtonStates();

                        // Display system information
                        DisplaySystemInfo();

                        // Display existing Devices and Terminals
                        ITcSmTreeItem devices = sysManager.LookupTreeItem("TIID");

                        if (devices == null)
                        {
                            Log("Could not find I/O Devices node.");
                            return;
                        }
                        EnumerateDevices(devices);
                    }
                    catch (Exception ex)
                    {
                        Log($"Error loading configuration: {ex.Message}");
                        MessageBox.Show($"Failed to load configuration.\n\n{ex.Message}",
                            "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void UpdateButtonStates()
        {
            btnLoadXTI.Enabled = configurationLoaded;
            btnScanNetwork.Enabled = configurationLoaded;
        }
        private void LoadXTIFile(string filePath)
        {
            try
            {
                Log($"Loading XTI file: {filePath}");

                // Import the XTI configuration
                sysManager.ConsumeMappingInfo(filePath);
                Log("XTI file loaded successfully.");

                // Display information about the loaded configuration
                DisplaySystemInfo();
                // Get the I/O Devices node
                ITcSmTreeItem devices = sysManager.LookupTreeItem("TIIC");

                if (devices == null)
                {
                    Log("Could not find I/O Devices node.");
                    return;
                }
                EnumerateDevices(devices);
            }
            catch (Exception ex)
            {
                Log($"Error loading XTI file: {ex.Message}");
                MessageBox.Show($"Failed to load XTI file.\n\n{ex.Message}",
                    "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnScanNetwork_Click(object sender, EventArgs e)
        {
            ScanForDevices();
        }

        private void ScanForDevices()
        {
            try
            {
                Log("Starting network scan for EtherCAT devices...");

                // Get the I/O Devices node
                ITcSmTreeItem devices = sysManager.LookupTreeItem("TIID");

                if (devices == null)
                {
                    Log("Could not find I/O Devices node.");
                    return;
                }

                // Scan for boxes (EtherCAT devices)
                Log("Scanning for devices...");
                ITcSmTreeItem scanResult = devices.CreateChild("Device", 0, "", null);

                // Perform scan
                scanResult.ConsumeXml("<TreeItem><DeviceGrp><ScanBoxes>1</ScanBoxes></DeviceGrp></TreeItem>");

                Log("Scan completed. Enumerating found devices...");

                // Enumerate found devices
                EnumerateDevices(devices);
            }
            catch (Exception ex)
            {
                Log($"Error scanning network: {ex.Message}");
                MessageBox.Show($"Failed to scan network.\n\n{ex.Message}",
                    "Scan Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EnumerateDevices(ITcSmTreeItem parent)
        {
            try
            {
                foreach (ITcSmTreeItem device in parent)
                {
                    Log($"  - {device.Name} (Type: {device.ItemType})");
                    AddDeviceCheckBox(device, device.Name, device.ItemType);
                    foreach (ITcSmTreeItem child in device)
                    {
                        if (child.ItemType == 5 || child.ItemType == 6)
                        {
                            Log($"    - {child.Name} (Type: {child.ItemType})");
                            AddDeviceCheckBox(child, child.Name, child.ItemType);
                        }   
                        foreach (ITcSmTreeItem child2 in child)
                        {
                            if (child2.ItemType == 5 || child2.ItemType == 6)
                            {
                                Log($"    - {child2.Name} (Type: {child2.ItemType})");
                                AddDeviceCheckBox(child2, child2.Name, child2.ItemType);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Error enumerating devices: {ex.Message}");
            }
        }

        private void AddDeviceCheckBox(ITcSmTreeItem device, string name, int type)
        {
            if (flowDevices.InvokeRequired)
            {
                flowDevices.Invoke(new Action(() => AddDeviceCheckBox(device, name, type)));
                return;
            }

            try
            {
                // Get current disabled state from device
                bool isDisabled = false;
                try
                {
                    string xml = device.ProduceXml(false);
                    // Parse XML to check if device is disabled
                    // This is a simplified check - you may need to adjust based on actual XML structure
                    isDisabled = xml.Contains("<Disabled>true</Disabled>") ||
                                 xml.Contains("Disabled=\"true\"");
                }
                catch { }

                CheckBox chk = new CheckBox
                {
                    Text = $"{name} ({type})",
                    AutoSize = true,
                    Checked = !isDisabled, // Checked means enabled
                    Margin = new Padding(5),
                    Width = 280
                };

                chk.CheckedChanged += (s, e) => DeviceCheckBox_CheckedChanged(chk, device);

                deviceCheckBoxes[chk] = device;
                flowDevices.Controls.Add(chk);
            }
            catch (Exception ex)
            {
                Log($"Error adding checkbox for device {name}: {ex.Message}");
            }
        }

        private void DeviceCheckBox_CheckedChanged(CheckBox checkBox, ITcSmTreeItem device)
        {
            try
            {
                bool shouldDisable = !checkBox.Checked;

                Log($"Setting device '{checkBox.Text}' to {(shouldDisable ? "DISABLED" : "ENABLED")}");

                // Get current XML configuration
                string currentXml = device.ProduceXml(false);

                // Modify the XML to set disabled state
                // This is a simplified approach - you may need to parse and modify XML more carefully
                string modifiedXml;
                if (currentXml.Contains("<Disabled>"))
                {
                    modifiedXml = System.Text.RegularExpressions.Regex.Replace(
                        currentXml,
                        "<Disabled>.*?</Disabled>",
                        $"<Disabled>{shouldDisable.ToString().ToLower()}</Disabled>");
                }
                else if (currentXml.Contains("Disabled="))
                {
                    modifiedXml = System.Text.RegularExpressions.Regex.Replace(
                        currentXml,
                        "Disabled=\".*?\"",
                        $"Disabled=\"{shouldDisable.ToString().ToLower()}\"");
                }
                else
                {
                    // Add disabled attribute if it doesn't exist
                    // This approach may need adjustment based on XML structure
                    modifiedXml = currentXml;
                }

                // Apply the modified configuration
                device.ConsumeXml(modifiedXml);

                Log($"Device state updated successfully.");
            }
            catch (Exception ex)
            {
                Log($"Error setting device state: {ex.Message}");
                MessageBox.Show($"Failed to change device state.\n\n{ex.Message}",
                    "Device Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Revert checkbox state
                checkBox.CheckedChanged -= (s, e) => DeviceCheckBox_CheckedChanged(checkBox, device);
                checkBox.Checked = !checkBox.Checked;
                checkBox.CheckedChanged += (s, e) => DeviceCheckBox_CheckedChanged(checkBox, device);
            }
        }

        private void ClearDeviceCheckBoxes()
        {
            if (flowDevices.InvokeRequired)
            {
                flowDevices.Invoke(new Action(ClearDeviceCheckBoxes));
                return;
            }

            flowDevices.Controls.Clear();
            deviceCheckBoxes.Clear();
        }
        private void DisplaySystemInfo()
        {
            try
            {
                Log("\n=== System Information ===");

                // Get system information
                ITcSmTreeItem system = sysManager.LookupTreeItem("TIRS");

                if (system != null)
                {
                    Log($"System Name: {system.Name}");

                    // You can extract more information from the XML
                    string systemXml = system.ProduceXml(false);
                    // Parse systemXml as needed for additional details
                }

                Log("=========================\n");
            }
            catch (Exception ex)
            {
                Log($"Error getting system info: {ex.Message}");
            }
        }

        private void Log(string message)
        {
            if (txtOutput.InvokeRequired)
            {
                txtOutput.Invoke(new Action(() => Log(message)));
            }
            else
            {
                txtOutput.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
                txtOutput.ScrollToCaret();
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (sysManager != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sysManager);
                    sysManager = null;
                }
            }
            catch (Exception ex)
            {
                Log($"Error during cleanup: {ex.Message}");
            }
        }
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}