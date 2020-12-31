using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GTAV_Firewall_Switcher
{
    internal enum RuleStatus
    {
        Missing,
        Enabled,
        Disabled
    }

    internal class TrayApplicationContext : ApplicationContext
    {
        private static readonly NotifyIcon TrayIcon = new NotifyIcon();
        private const string Program = "GTA5";
        private const string Rule = "GTAVFirewallSwitcherRule" + Program;
        private readonly KeyboardHook _hook = new KeyboardHook();

        public TrayApplicationContext()
        {
            TrayIcon.Icon = Properties.Resources.AppIcon;
            TrayIcon.Text = "Icon";
            var contextMenuStrip = new ContextMenuStrip();
            ToolStripItem infoItem = new ToolStripMenuItem("Ctrl + Alt + F12 to toggle");
            infoItem.Enabled = false;
            ToolStripItem exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += Exit;
            ToolStripItem blockItem = new ToolStripMenuItem("Block");
            blockItem.Click += Block;
            ToolStripItem allowItem = new ToolStripMenuItem("Allow");
            allowItem.Click += Allow;
            contextMenuStrip.Items.AddRange(new[] {infoItem, blockItem, allowItem, exitItem});

            TrayIcon.ContextMenuStrip = contextMenuStrip;
            TrayIcon.Visible = true;

            _hook.RegisterHotKey(ModifierKeys.Control | ModifierKeys.Alt, Keys.F12);
            _hook.KeyPressed += (_, args) => ToggleRule(args);
        }

        private static void ToggleRule(KeyPressedEventArgs e)
        {
            var result = CreateRuleIfNotExists();
            var success = result switch
            {
                RuleStatus.Missing => false,
                RuleStatus.Enabled => DisableRule(),
                RuleStatus.Disabled => EnableRule(),
                _ => false,
            };
            Console.WriteLine("Toggle Rule success: " + success);
        }

        private static void Block(object sender, EventArgs e)
        {
            var result = CreateRuleIfNotExists();
            var success = result switch
            {
                // Result is .Missing if the process was not found
                RuleStatus.Missing => false,
                RuleStatus.Disabled => EnableRule(),
                _ => true
            };
            Console.WriteLine("Block: " + success);
        }

        private static void Allow(object sender, EventArgs e)
        {
            var result = CreateRuleIfNotExists();
            var success = result switch
            {
                // Result is .Missing if the process was not found
                RuleStatus.Missing => false,
                RuleStatus.Enabled => DisableRule(),
                _ => true
            };
            Console.WriteLine("Allow: " + success);
        }

        private static void Speak(string textToSpeech, bool wait = false)
        {
            // Command to execute PS  
            Execute($@"Add-Type -AssemblyName System.speech;  
            $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer;                           
            $speak.Speak(""{textToSpeech}"");"); // Embedd text  

            void Execute(string command)
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
        }

        private static RuleStatus CreateRuleIfNotExists()
        {
            var process = Process.GetProcessesByName(Program);
            if (process.Length < 1 || process[0].MainModule == null)
                return RuleStatus.Missing;

            var ruleStatus = GetRuleStatus();
            if (ruleStatus != RuleStatus.Missing) return ruleStatus;

            CreateRule(process[0].MainModule.FileName);
            return RuleStatus.Disabled;
        }

        private static bool CreateRule(string programPath)
        {
            var output1 = RunCommand("netsh",
                $"advfirewall firewall add rule program=\"{programPath}\" name=\"{Rule}\" enable=no dir=out action=block");
            var output2 = RunCommand("netsh",
                $"advfirewall firewall add rule program=\"{programPath}\" name=\"{Rule}\" enable=no dir=in action=block");
            Console.WriteLine("Create Rule:\n" + output1 + "\n\n" + output2);
            return output1.Contains("Ok.") && output2.Contains("Ok.");
        }

        private static bool EnableRule()
        {
            var output = RunCommand("netsh", $"advfirewall firewall set rule name=\"{Rule}\" new enable=yes");
            Console.WriteLine("EnableRule");
            var success = output.Contains("Ok.");
            Speak(success ? "Enable firewall rule" : "Disabling rule failed");
            return success;
        }

        private static bool DisableRule()
        {
            var output = RunCommand("netsh", $"advfirewall firewall set rule name=\"{Rule}\" new enable=no");
            Console.WriteLine("DisableRule");
            var success = output.Contains("Ok.");
            Speak(success ? "Disabled firewall rule" : "Disabling rule failed");
            return success;
        }

        private static RuleStatus GetRuleStatus()
        {
            var output = RunCommand("netsh", $"advfirewall firewall show rule name=\"{Rule}\"");
            var exists = !output.Contains("No rules match the specified criteria") && output.Contains("Rule Name:") &&
                         output.Contains("Enabled:");
            if (!exists)
                return RuleStatus.Missing;
            var enabledLine = output.Split(Environment.NewLine).First(line => line.Contains("Enabled:"));

            return enabledLine.EndsWith("No") ? RuleStatus.Disabled : RuleStatus.Enabled;
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
                }
            };
            process.Start();
            return process.StandardOutput.ReadToEnd();
        }

        private static void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            TrayIcon.Visible = false;
            Application.Exit();
        }
    }
}