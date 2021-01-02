using System;
using System.Diagnostics;
using System.Text;
using System.Net.NetworkInformation;
using System.Windows.Forms;

namespace InternetSwitcher
{
    internal class TrayApplicationContext : ApplicationContext
    {
        private static readonly NotifyIcon TrayIcon = new NotifyIcon();
        private readonly KeyboardHook _hook = new KeyboardHook();

        public TrayApplicationContext()
        {
            TrayIcon.Icon = Properties.Resources.AppIcon;
            TrayIcon.Text = "Internet Switcher";
            var contextMenuStrip = new ContextMenuStrip();
            ToolStripItem infoItem = new ToolStripMenuItem("Ctrl + Alt + F12 to toggle");
            infoItem.Enabled = false;
            ToolStripItem exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += Exit;
            ToolStripItem blockItem = new ToolStripMenuItem("Enable Internet");
            blockItem.Click += (a, b) => EnableInternet();
            ToolStripItem allowItem = new ToolStripMenuItem("Disable Internet");
            allowItem.Click += (a, b) => DisableInternet();
            contextMenuStrip.Items.AddRange(new[] {infoItem, blockItem, allowItem, exitItem});

            TrayIcon.ContextMenuStrip = contextMenuStrip;
            TrayIcon.Visible = true;

            _hook.RegisterHotKey(ModifierKeys.Control | ModifierKeys.Alt, Keys.F12);
            _hook.KeyPressed += (a, b) => ToggleRule();
        }

        private static void ToggleRule()
        {
            if (IsInternetConnected())
            {
                DisableInternet();
            }
            else
            {
                EnableInternet();
            }
        }

        private static void DisableInternet()
        {
            Speak("Disabling...");
            RunCommand("ipconfig", "/release");
            Speak("Disabled internet connection");
        }

        private static void EnableInternet()
        {
            Speak("Enabling...");
            RunCommand("ipconfig", "/renew");
            Speak("Enabled internet connection");
        }

        private static bool IsInternetConnected()
        {
            try
            {
                var myPing = new Ping();
                const string host = "google.com";
                var buffer = new byte[32];
                const int timeout = 1000;
                var pingOptions = new PingOptions();
                var reply = myPing.Send(host, timeout, buffer, pingOptions);
                return reply != null && reply.Status == IPStatus.Success;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            TrayIcon.Visible = false;
            Application.Exit();
        }

        private static void Speak(string textToSpeech, bool wait = false)
        {
            // Command to execute PS  
            PowerShell($@"Add-Type -AssemblyName System.speech;  
            $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer;                           
            $speak.Speak(""{textToSpeech}"");", wait); // Embedd text  
        }

        private static void PowerShell(string command, bool wait = true)
        {
            // create a temp file with .ps1 extension  
            var cFile = System.IO.Path.GetTempPath() + Guid.NewGuid() + ".ps1";

            //Write the .ps1  
            using var tw = new System.IO.StreamWriter(cFile, false, Encoding.UTF8);
            tw.Write(command);

            // Setup the PS  
            var start =
                new ProcessStartInfo
                {
                    FileName =
                        "C:\\windows\\system32\\windowspowershell\\v1.0\\powershell.exe", // CHUPA MICROSOFT 02-10-2019 23:45                    
                    LoadUserProfile = false,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Arguments = $"-executionpolicy bypass -File {cFile}",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

            //Init the Process  
            var p = Process.Start(start);
            // The wait may not work! :(  
            if (wait) p?.WaitForExit();
        }

        private static string RunCommand(string file, string arguments)
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    WindowStyle = ProcessWindowStyle.Hidden,
                    FileName = file,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                }
            };
            process.Start();
            return process.StandardOutput.ReadToEnd();
        }
    }
}