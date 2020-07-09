using mp.hooks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using System.Management.Automation;
using System.IO;

using VirtualDesktop;

namespace MyHotKeys
{
    public partial class Form1 : Form
    {
        const string PATH_PS = "C:\\ps\\";

        protected static string userDir = Environment.GetEnvironmentVariable("USERPROFILE");

        protected string pathWallpapers = userDir + @"\wallpapers\";
        protected static string pathPsxIsos = userDir + @"\Downloads\psx\";
        protected string pathPcsx2Isos = pathPsxIsos + @"psx2\";

        protected string pathExePsx = userDir + @"\programs\epsxe202-1\epsxe.exe";        
        protected string pathExePcsx2 = userDir + @"\programs\PCSX2 1.6.0\pcsx2.exe";
        protected string pathExeOpera = userDir + @"\AppData\Local\Programs\Opera developer\launcher.exe";

        protected List<string> psxIsos = new List<string>() {
            "FIFA Soccer 2005",
            "Crash Bandicoot",
            "CTR - Crash Team Racing",
            "Need for Speed III - Hot Pursuit",
            "Hogs of War",
            "Grand Theft Auto - Liberty City Stories",
            "Sled Storm"
        };

        protected Dictionary<string, string> multiPsxIsos = new Dictionary<string, string> {
            ["Hogs of War"] = "Hogs of War (Track 1)",
            ["Sled Storm"] = "Sled Storm (Track 01)"
        };

        protected Dictionary<string, int> psVersion = new Dictionary<string, int> {
            ["Grand Theft Auto - Liberty City Stories"] = 2
        };

        int[] acceptWorkspace = {0, 1, 2, 3};

        KeyboardHook kh;
        bool bCloseFromTray = false;
        int idCurrentWorkspace = 1;

        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;

        public Form1()
        {
            InitializeComponent();

            PreparePaths();

            SetHook();

            CreateNotifyIcon();
            CheckWorkspace();
        }

        protected void PreparePaths()
        {
            string path = Directory.GetParent(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            ).FullName;

            if (Environment.OSVersion.Version.Major >= 6)
            {
                this.pathWallpapers = Directory.GetParent(path).ToString() + "\\wallpapers\\";
            }
        }

        protected void SetHook()
        {
            kh = new KeyboardHook();

            kh.hook();
            kh.KeyUp += KeyP;
        }

        private void KeyP(int wParam, KeyboardHookData lParam)
        {
            if (kh.WinHeld && kh.CtrlHeld)
            {
                if (kh.AltHeld && lParam.vkCode != 115) { // Alt, but not F4
                    if (lParam.vkCode == 121) // F10
                    {
                        RunPsx();
                    }
                    else if (lParam.vkCode == 120) // F9
                    {
                        WinRun();
                    }
                    else if (lParam.vkCode == 119) // F8
                    {
                        OperaPrivate();
                    }
                }
                else if (lParam.vkCode > 48 && lParam.vkCode <= 52 && this.acceptWorkspace.Contains(lParam.vkCode - 49))
                { // 1-4
                    MoveToWorkspace(lParam.vkCode - 49);
                }
                else if (lParam.vkCode == 48) // 0
                {
                    RunPsx(-2);
                }
                else if (lParam.vkCode == 56) // 8
                {
                    RunPsx(0);
                }
                else if (lParam.vkCode > 53 && lParam.vkCode < 58) // 6-9
                {
                    RunPsx(lParam.vkCode - 60);
                }
                else if (lParam.vkCode >= 122 && lParam.vkCode <= 123) // F11-F12
                {
                    RunPsx(lParam.vkCode - 123);
                }
                else if (lParam.vkCode >= 112 && lParam.vkCode <= 123) // F1-F12
                {
                    ChangeBackground(lParam.vkCode - 111);
                }
                else if (lParam.vkCode == 37 && this.acceptWorkspace.Contains(this.idCurrentWorkspace - 1))
                { // Left
                     this.idCurrentWorkspace--;
                     SetTrayIconN(this.idCurrentWorkspace);
                }
                else if (lParam.vkCode == 39 && this.acceptWorkspace.Contains(this.idCurrentWorkspace + 1))
                { // Right
                     this.idCurrentWorkspace++;
                     SetTrayIconN(this.idCurrentWorkspace);
                }
            }
        }

        [DllImport("kernel32.dll")]
        static extern uint WinExec(string lpCmdLine, uint uCmdShow);
        protected void OperaPrivate()
        {
            WinExec(pathExeOpera + " --private", 0);
        }

        protected void RunPsx()
        {
            WinExec(pathExePsx, 1);
        }
        protected void RunPsx(int revertIdPsxIso)
        {
            string isoName = psxIsos[-revertIdPsxIso];
            string multiIsoName = multiPsxIsos.ContainsKey(isoName) ? multiPsxIsos[isoName] : isoName;

            if (psVersion.ContainsKey(isoName) && psVersion[isoName] == 2) {
                RunPcsx2(isoName, multiIsoName);
                return;
            }

            WinExec(pathExePsx + $@" -nogui -loadbin ""{pathPsxIsos}/{isoName}/{multiIsoName}.bin""", 1);
        }

        protected void RunPcsx2(string isoName, string multiIsoName = null)
        {
            if (multiIsoName == null)
                multiIsoName = isoName;

            WinExec(pathExePcsx2 + $@" --nogui --fullscreen ""{pathPcsx2Isos}/{isoName}/{multiIsoName}.iso""", 1);
        }

        protected void MoveToWorkspace(int idWorkspace)
        {
            PowerShell ps = PowerShell.Create();

            ps.AddScript(PATH_PS + "vd" + (idWorkspace + 1).ToString() + ".ps1");
            ps.Invoke();

            //Desktop.FromIndex(idWorkspace).MakeVisible();

            this.idCurrentWorkspace = idWorkspace;
            SetTrayIconN(idWorkspace);
        }

        [DllImport("user32.dll")]
        static extern IntPtr SystemParametersInfo(int uiAction, int uiParam, string path, int fWinIni);
        protected void ChangeBackground(int idWallpaper)
        {
            SystemParametersInfo(0x0014, 0, this.pathWallpapers + "bg" + idWallpaper.ToString() + ".png", 0x0001);
        }

        protected void WinRun()
        {
            var dialogBox = new Form();
            var textBox = new TextBox();
            var buttonOk = new Button();
            var buttonCancel = new Button();

            dialogBox.Size = new System.Drawing.Size(400, 100);
            dialogBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialogBox.StartPosition = FormStartPosition.CenterScreen;
            dialogBox.Text = "Run...";            

            textBox.Width = 365;
            textBox.Location = new Point(10, 10);

            buttonOk.Text = "Run";
            buttonOk.Location = new Point(220, 35);
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.Text = "Cancel";
            buttonCancel.Location = new Point(300, 35);
            buttonCancel.DialogResult = DialogResult.Cancel;

            dialogBox.AcceptButton = buttonOk;
            dialogBox.CancelButton = buttonCancel;
            dialogBox.Controls.Add(textBox);
            dialogBox.Controls.Add(buttonOk);
            dialogBox.Controls.Add(buttonCancel);
            dialogBox.ShowDialog();            

            if (dialogBox.DialogResult == DialogResult.OK)
            {
                WinExec(textBox.Text, 1);
            }

            dialogBox.Dispose();
        }

        private void CreateNotifyIcon()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.menuItem1 = new System.Windows.Forms.MenuItem();

            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.notifyIcon1.Icon = new System.Drawing.Icon("icon.ico");
            this.notifyIcon1.ContextMenu = this.contextMenu1;
            this.notifyIcon1.Text = "My Hot Keys";
            this.notifyIcon1.Visible = true;

            this.menuItem1.Text = "E&xit";
            this.menuItem1.Index = 0;
            this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);

            this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
            {
                this.menuItem1
            });

            this.notifyIcon1.Click += new System.EventHandler(this.notifyIcon1_Click);
        }

        private void CheckWorkspace()
        {
            this.idCurrentWorkspace = Desktop.FromDesktop(Desktop.Current);
            SetTrayIconN(this.idCurrentWorkspace);
        }

        private void SetTrayIconN(int num)
        {
            string fileName = "icon" + num.ToString() + ".ico";

            if (this.acceptWorkspace.Contains(num) && File.Exists(fileName))
            {
                this.notifyIcon1.Icon = new System.Drawing.Icon(fileName);
            }            
        }

        private void notifyIcon1_Click(object Sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }

        private void menuItem1_Click(object Sender, EventArgs e)
        {
            this.bCloseFromTray = true;
            Application.Exit();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            WindowState = FormWindowState.Minimized;
            e.Cancel = !bCloseFromTray;
        }
    }
}
