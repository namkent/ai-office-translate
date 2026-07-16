using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;

namespace AITranslateVSTOInstaller
{
    public class SetupForm : Form
    {
        private Panel pnlTitleBar;
        private PictureBox picIcon;
        private Label lblTitleText;
        private Button btnTitleClose;

        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblStatus;
        private ProgressBar progressBar;
        private Button btnInstall;
        private Button btnUninstall;
        private Button btnClose;

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        private void TitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        public SetupForm()
        {
            LoadAppIcon();
            InitializeComponent();
        }

        private void LoadAppIcon()
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string actualName = null;
                foreach (string name in assembly.GetManifestResourceNames())
                {
                    if (name.EndsWith("icon.ico", StringComparison.OrdinalIgnoreCase))
                    {
                        actualName = name;
                        break;
                    }
                }
                if (actualName != null)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(actualName))
                    {
                        if (stream != null)
                        {
                            this.Icon = new Icon(stream);
                        }
                    }
                }
            }
            catch { }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            LoadTitleBarIcon();
        }

        private void LoadTitleBarIcon()
        {
            bool loaded = false;
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string pngResourceName = null;
                foreach (string name in assembly.GetManifestResourceNames())
                {
                    if (name.EndsWith("icon.png", StringComparison.OrdinalIgnoreCase))
                    {
                        pngResourceName = name;
                        break;
                    }
                }
                if (pngResourceName != null)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(pngResourceName))
                    {
                        if (stream != null)
                        {
                            picIcon.Image = Image.FromStream(stream);
                            loaded = true;
                        }
                    }
                }
            }
            catch { }

            if (!loaded)
            {
                picIcon.Visible = false;
                lblTitleText.Left = 10;
            }
        }

        private void InitializeComponent()
        {
            this.Text = "AI Translate VSTO Installer";
            this.ClientSize = new Size(480, 227);
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(243, 243, 243);

            // Custom Title Bar Panel (Màu sáng)
            pnlTitleBar = new Panel();
            pnlTitleBar.Location = new Point(0, 0);
            pnlTitleBar.Size = new Size(480, 32);
            pnlTitleBar.BackColor = Color.FromArgb(230, 230, 230);
            pnlTitleBar.MouseDown += new MouseEventHandler(TitleBar_MouseDown);

            // Icon PictureBox
            picIcon = new PictureBox();
            picIcon.Location = new Point(10, 8);
            picIcon.Size = new Size(16, 16);
            picIcon.SizeMode = PictureBoxSizeMode.Zoom;

            // Title Text Label
            lblTitleText = new Label();
            lblTitleText.Text = "AI Translate VSTO Installer (No Admin)";
            lblTitleText.Font = new Font("Segoe UI", 9, FontStyle.Regular);
            lblTitleText.ForeColor = Color.FromArgb(27, 27, 27);
            lblTitleText.Location = new Point(32, 7);
            lblTitleText.Size = new Size(300, 20);
            lblTitleText.MouseDown += new MouseEventHandler(TitleBar_MouseDown);

            // Title Close Button
            btnTitleClose = new Button();
            btnTitleClose.Text = "✕";
            btnTitleClose.Font = new Font("Segoe UI", 9, FontStyle.Regular);
            btnTitleClose.BackColor = Color.FromArgb(230, 230, 230);
            btnTitleClose.ForeColor = Color.FromArgb(27, 27, 27);
            btnTitleClose.FlatStyle = FlatStyle.Flat;
            btnTitleClose.FlatAppearance.BorderSize = 0;
            btnTitleClose.TabStop = false;
            btnTitleClose.Location = new Point(434, 0);
            btnTitleClose.Size = new Size(46, 32);
            btnTitleClose.Cursor = Cursors.Hand;
            btnTitleClose.Click += (s, e) => this.Close();
            btnTitleClose.MouseEnter += (s, e) => {
                btnTitleClose.BackColor = Color.FromArgb(232, 17, 35);
                btnTitleClose.ForeColor = Color.White;
            };
            btnTitleClose.MouseLeave += (s, e) => {
                btnTitleClose.BackColor = pnlTitleBar.BackColor;
                btnTitleClose.ForeColor = Color.FromArgb(27, 27, 27);
            };

            pnlTitleBar.Controls.Add(picIcon);
            pnlTitleBar.Controls.Add(lblTitleText);
            pnlTitleBar.Controls.Add(btnTitleClose);

            // Title Label
            lblTitle = new Label();
            lblTitle.Text = "AI Office Translate (VSTO)";
            lblTitle.Font = new Font("Segoe UI Semibold", 16, FontStyle.Bold);
            lblTitle.ForeColor = Color.FromArgb(27, 27, 27);
            lblTitle.Location = new Point(30, 57);
            lblTitle.Size = new Size(420, 32);

            // Subtitle Label
            lblSubtitle = new Label();
            lblSubtitle.Text = "Modern VSTO Add-in deployment for Enterprise Office";
            lblSubtitle.Font = new Font("Segoe UI", 9, FontStyle.Regular);
            lblSubtitle.ForeColor = Color.FromArgb(95, 95, 95);
            lblSubtitle.Location = new Point(32, 89);
            lblSubtitle.Size = new Size(420, 20);

            // Status Label
            lblStatus = new Label();
            lblStatus.Text = "Ready to install or uninstall the VSTO add-in.";
            lblStatus.Font = new Font("Segoe UI", 9, FontStyle.Regular);
            lblStatus.ForeColor = Color.FromArgb(95, 95, 95);
            lblStatus.Location = new Point(32, 117);
            lblStatus.Size = new Size(420, 18);
            lblStatus.Visible = false;

            // Progress Bar
            progressBar = new ProgressBar();
            progressBar.Location = new Point(32, 140);
            progressBar.Size = new Size(420, 10);
            progressBar.Style = ProgressBarStyle.Blocks;
            progressBar.BackColor = Color.FromArgb(230, 230, 230);
            progressBar.Visible = false;

            // Install Button
            btnInstall = new Button();
            btnInstall.Text = "INSTALL";
            btnInstall.Font = new Font("Segoe UI Semibold", 9, FontStyle.Bold);
            btnInstall.BackColor = Color.FromArgb(0, 95, 184);
            btnInstall.ForeColor = Color.White;
            btnInstall.FlatStyle = FlatStyle.Flat;
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Location = new Point(30, 165);
            btnInstall.Size = new Size(120, 32);
            btnInstall.Cursor = Cursors.Hand;
            btnInstall.Click += new EventHandler(BtnInstall_Click);
            btnInstall.MouseEnter += (s, e) => { if (!isWorking) btnInstall.BackColor = Color.FromArgb(24, 115, 204); };
            btnInstall.MouseLeave += (s, e) => { if (!isWorking) btnInstall.BackColor = Color.FromArgb(0, 95, 184); };

            // Uninstall Button
            btnUninstall = new Button();
            btnUninstall.Text = "UNINSTALL";
            btnUninstall.Font = new Font("Segoe UI Semibold", 9, FontStyle.Bold);
            btnUninstall.BackColor = Color.FromArgb(253, 253, 253);
            btnUninstall.ForeColor = Color.FromArgb(27, 27, 27);
            btnUninstall.FlatStyle = FlatStyle.Flat;
            btnUninstall.FlatAppearance.BorderSize = 1;
            btnUninstall.FlatAppearance.BorderColor = Color.FromArgb(229, 229, 229);
            btnUninstall.Location = new Point(180, 165);
            btnUninstall.Size = new Size(120, 32);
            btnUninstall.Cursor = Cursors.Hand;
            btnUninstall.Click += new EventHandler(BtnUninstall_Click);
            btnUninstall.MouseEnter += (s, e) => { if (!isWorking) btnUninstall.BackColor = Color.FromArgb(245, 245, 245); };
            btnUninstall.MouseLeave += (s, e) => { if (!isWorking) btnUninstall.BackColor = Color.FromArgb(253, 253, 253); };

            // Close Button
            btnClose = new Button();
            btnClose.Text = "CLOSE";
            btnClose.Font = new Font("Segoe UI Semibold", 9, FontStyle.Bold);
            btnClose.BackColor = Color.FromArgb(253, 253, 253);
            btnClose.ForeColor = Color.FromArgb(27, 27, 27);
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.FlatAppearance.BorderSize = 1;
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(229, 229, 229);
            btnClose.Location = new Point(330, 165);
            btnClose.Size = new Size(120, 32);
            btnClose.Cursor = Cursors.Hand;
            btnClose.Click += (s, e) => this.Close();
            btnClose.MouseEnter += (s, e) => { if (!isWorking) btnClose.BackColor = Color.FromArgb(245, 245, 245); };
            btnClose.MouseLeave += (s, e) => { if (!isWorking) btnClose.BackColor = Color.FromArgb(253, 253, 253); };

            // Add Controls
            this.Controls.Add(pnlTitleBar);
            this.Controls.Add(lblTitle);
            this.Controls.Add(lblSubtitle);
            this.Controls.Add(lblStatus);
            this.Controls.Add(progressBar);
            this.Controls.Add(btnInstall);
            this.Controls.Add(btnUninstall);
            this.Controls.Add(btnClose);
        }

        private void SetStatus(string message, int progressPercent, bool isMarquee = false)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string, int, bool>(SetStatus), message, progressPercent, isMarquee);
                return;
            }
            lblStatus.Text = message;
            progressBar.Style = isMarquee ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
            if (!isMarquee)
            {
                progressBar.Value = progressPercent;
            }
        }

        private bool isWorking = false;

        private void SetControlsEnabled(bool enabled)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool>(SetControlsEnabled), enabled);
                return;
            }
            isWorking = !enabled;
            var cursor = enabled ? Cursors.Hand : Cursors.WaitCursor;
            btnInstall.Cursor = cursor;
            btnUninstall.Cursor = cursor;
            btnClose.Cursor = cursor;
        }

        private void BtnInstall_Click(object sender, EventArgs e)
        {
            if (isWorking) return;
            lblStatus.Visible = true;
            progressBar.Visible = true;
            SetControlsEnabled(false);
            Thread thread = new Thread(DoInstall);
            thread.Start();
        }

        private void BtnUninstall_Click(object sender, EventArgs e)
        {
            if (isWorking) return;
            lblStatus.Visible = true;
            progressBar.Visible = true;
            SetControlsEnabled(false);
            Thread thread = new Thread(DoUninstall);
            thread.Start();
        }

        private void DoInstall()
        {
            try
            {
                SetStatus("Stopping Excel, Word, and PowerPoint...", 10, true);
                KillOfficeProcesses();
                Thread.Sleep(1000);

                SetStatus("Preparing target directory...", 25, false);
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string targetDir = Path.Combine(appData, "AITranslateVSTO");
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                SetStatus("Extracting VSTO binary packages...", 50, false);
                // Extract PFX
                string pfxPath = Path.Combine(targetDir, "AITranslate.pfx");
                ExtractResource("AITranslate.pfx", pfxPath);

                // Extract Excel Add-in
                string excelDir = Path.Combine(targetDir, "AITranslateExcel");
                if (!Directory.Exists(excelDir)) Directory.CreateDirectory(excelDir);
                ExtractResource("AITranslateExcel.dll", Path.Combine(excelDir, "AITranslateExcel.dll"));
                ExtractResource("AITranslateExcel.dll.manifest", Path.Combine(excelDir, "AITranslateExcel.dll.manifest"));
                ExtractResource("AITranslateExcel.vsto", Path.Combine(excelDir, "AITranslateExcel.vsto"));

                // Extract Word Add-in
                string wordDir = Path.Combine(targetDir, "AITranslateWord");
                if (!Directory.Exists(wordDir)) Directory.CreateDirectory(wordDir);
                ExtractResource("AITranslateWord.dll", Path.Combine(wordDir, "AITranslateWord.dll"));
                ExtractResource("AITranslateWord.dll.manifest", Path.Combine(wordDir, "AITranslateWord.dll.manifest"));
                ExtractResource("AITranslateWord.vsto", Path.Combine(wordDir, "AITranslateWord.vsto"));

                // Extract PowerPoint Add-in
                string pptDir = Path.Combine(targetDir, "AITranslatePPT");
                if (!Directory.Exists(pptDir)) Directory.CreateDirectory(pptDir);
                ExtractResource("AITranslatePPT.dll", Path.Combine(pptDir, "AITranslatePPT.dll"));
                ExtractResource("AITranslatePPT.dll.manifest", Path.Combine(pptDir, "AITranslatePPT.dll.manifest"));
                ExtractResource("AITranslatePPT.vsto", Path.Combine(pptDir, "AITranslatePPT.vsto"));

                SetStatus("Importing self-signed certificate (Trusting)...", 75, true);
                TrustCertificate(pfxPath);

                SetStatus("Writing settings configuration...", 85, false);
                WriteSettings(targetDir);

                SetStatus("Registering VSTO manifests in registry...", 90, false);
                ConfigureRegistry(targetDir);

                SetStatus("Installation Successful!", 100, false);
                SetControlsEnabled(true);
                this.Invoke(new Action(() => {
                    MessageBox.Show(this, "AI Office Translate VSTO Add-ins have been successfully installed!\nPlease restart Excel, Word, and PowerPoint.", "Installation Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            }
            catch (Exception ex)
            {
                SetStatus("Installation Failed!", 0, false);
                SetControlsEnabled(true);
                this.Invoke(new Action(() => {
                    MessageBox.Show(this, "An error occurred during installation:\n\n" + ex.Message, "Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
            finally
            {
                SetControlsEnabled(true);
            }
        }

        private void DoUninstall()
        {
            try
            {
                SetStatus("Stopping Excel, Word, and PowerPoint...", 20, true);
                KillOfficeProcesses();
                Thread.Sleep(1000);

                SetStatus("Removing registry keys...", 50, false);
                CleanRegistry();

                SetStatus("Deleting target directory and files...", 80, false);
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string targetDir = Path.Combine(appData, "AITranslateVSTO");
                if (Directory.Exists(targetDir))
                {
                    try
                    {
                        Directory.Delete(targetDir, true);
                    }
                    catch (IOException)
                    {
                        Thread.Sleep(1000);
                        Directory.Delete(targetDir, true);
                    }
                }

                SetStatus("Uninstall Successful!", 100, false);
                SetControlsEnabled(true);
                this.Invoke(new Action(() => {
                    MessageBox.Show(this, "AI Office Translate VSTO Add-ins have been successfully uninstalled.", "Uninstall Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            }
            catch (Exception ex)
            {
                SetStatus("Uninstall Failed!", 0, false);
                SetControlsEnabled(true);
                this.Invoke(new Action(() => {
                    MessageBox.Show(this, "An error occurred during uninstallation:\n\n" + ex.Message, "Uninstall Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
            finally
            {
                SetControlsEnabled(true);
            }
        }

        private void KillOfficeProcesses()
        {
            string[] names = { "excel", "winword", "powerpnt" };
            foreach (string name in names)
            {
                Process[] processes = Process.GetProcessesByName(name);
                foreach (Process p in processes)
                {
                    try
                    {
                        p.Kill();
                        p.WaitForExit(5000);
                    }
                    catch { }
                }
            }
        }

        private void ExtractResource(string resourceName, string targetPath)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string actualName = null;
            foreach (string name in assembly.GetManifestResourceNames())
            {
                if (name.EndsWith(resourceName, StringComparison.OrdinalIgnoreCase))
                {
                    actualName = name;
                    break;
                }
            }

            if (actualName == null)
            {
                throw new Exception("Embedded resource not found: " + resourceName);
            }

            using (Stream stream = assembly.GetManifestResourceStream(actualName))
            using (FileStream fs = new FileStream(targetPath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fs);
            }
        }

        private void TrustCertificate(string pfxPath)
        {
            try
            {
                X509Certificate2 cert = new X509Certificate2(pfxPath, "secure-password-123", X509KeyStorageFlags.PersistKeySet);

                // Add to Trusted Publishers (CurrentUser - No Admin Required)
                using (X509Store store = new X509Store(StoreName.TrustedPublisher, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadWrite);
                    store.Add(cert);
                }

                // Add to Trusted Root (CurrentUser - No Admin Required)
                using (X509Store store = new X509Store(StoreName.Root, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadWrite);
                    store.Add(cert);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to import signing certificate: " + ex.Message);
            }
        }

        private void WriteSettings(string targetDir)
        {
            string apiUrl = "https://localhost:3000";
            string token = "secure-token-123";

            // Try to read .env from setup.exe directory
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            string envPath = Path.Combine(currentDir, ".env");
            if (!File.Exists(envPath))
            {
                envPath = Path.Combine(Path.GetDirectoryName(currentDir.TrimEnd('\\')), ".env");
            }

            if (File.Exists(envPath))
            {
                try
                {
                    string[] lines = File.ReadAllLines(envPath);
                    foreach (string line in lines)
                    {
                        if (line.Contains("="))
                        {
                            int idx = line.IndexOf('=');
                            string key = line.Substring(0, idx).Trim();
                            string val = line.Substring(idx + 1).Trim();

                            if (key == "PORT")
                            {
                                apiUrl = "https://localhost:" + val;
                            }
                            else if (key == "CLIENT_TOKEN" || key == "CLIENT_ACCESS_TOKEN")
                            {
                                token = val;
                            }
                        }
                    }
                }
                catch { }
            }

            // Save settings to common path expected by Core
            string settingsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AITranslateAddin");
            if (!Directory.Exists(settingsDir)) Directory.CreateDirectory(settingsDir);

            File.WriteAllLines(Path.Combine(settingsDir, "settings.txt"), new string[] {
                "API_URL=" + apiUrl,
                "TOKEN=" + token
            });
        }

        private void ConfigureRegistry(string targetDir)
        {
            var apps = new[] {
                new { Name = "Excel", Proj = "AITranslateExcel", RegName = "AITranslate.Excel", Desc = "Excel Translation Add-in using AI" },
                new { Name = "Word", Proj = "AITranslateWord", RegName = "AITranslate.Word", Desc = "Word Translation Add-in using AI" },
                new { Name = "PowerPoint", Proj = "AITranslatePPT", RegName = "AITranslate.PPT", Desc = "PowerPoint Translation Add-in using AI" }
            };

            foreach (var app in apps)
            {
                string vstoFile = Path.Combine(targetDir, app.Proj, app.Proj + ".vsto");
                string manifestValue = "file:///" + vstoFile.Replace('\\', '/') + "|vstolocal";
                string regPath = @"Software\Microsoft\Office\" + app.Name + @"\Addins\" + app.RegName;

                using (RegistryKey rk = Registry.CurrentUser.CreateSubKey(regPath))
                {
                    rk.SetValue("FriendlyName", "AI Office Translate");
                    rk.SetValue("Description", app.Desc);
                    rk.SetValue("Manifest", manifestValue);
                    rk.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }
            }
        }

        private void CleanRegistry()
        {
            string[] paths = {
                @"Software\Microsoft\Office\Excel\Addins\AITranslate.Excel",
                @"Software\Microsoft\Office\Word\Addins\AITranslate.Word",
                @"Software\Microsoft\Office\PowerPoint\Addins\AITranslate.PPT"
            };

            foreach (string path in paths)
            {
                try
                {
                    Registry.CurrentUser.DeleteSubKeyTree(path, false);
                }
                catch { }
            }

            // Also clean settings.txt folder
            try
            {
                string settingsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AITranslateAddin");
                if (Directory.Exists(settingsDir))
                {
                    Directory.Delete(settingsDir, true);
                }
            }
            catch { }
        }
    }

    public static class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SetupForm());
        }
    }
}
