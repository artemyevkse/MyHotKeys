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

namespace MyHotKeys
{
    public partial class Form1 : Form
    {
        const string PATH_PS = "C:\\ps\\";

        protected string pathWallpapers = "C:\\Users\\Константин\\wallpapers\\";

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
                if (lParam.vkCode > 48 && lParam.vkCode <= 52
                    && this.acceptWorkspace.Contains(lParam.vkCode - 49))
                {
                    MoveToWorkspace(lParam.vkCode - 49);
                }
                else if (lParam.vkCode >= 112 && lParam.vkCode <= 123)
                {
                    ChangeBackground(lParam.vkCode - 111);
                }
                else if (lParam.vkCode == 37
                    && this.acceptWorkspace.Contains(this.idCurrentWorkspace - 1))
                {
                     this.idCurrentWorkspace--;
                     SetTrayIconN(this.idCurrentWorkspace);
                }
                else if (lParam.vkCode == 39
                    && this.acceptWorkspace.Contains(this.idCurrentWorkspace + 1))
                {
                     this.idCurrentWorkspace++;
                     SetTrayIconN(this.idCurrentWorkspace);
                }
            }
        }

        protected void MoveToWorkspace(int idWorkspace)
        {
            PowerShell ps = PowerShell.Create();

            ps.AddScript(PATH_PS + "vd" + (idWorkspace + 1).ToString() + ".ps1");
            ps.Invoke();

            this.idCurrentWorkspace = idWorkspace;
            SetTrayIconN(idWorkspace);
        }

        [DllImport("user32.dll")]
        static extern IntPtr SystemParametersInfo(int uiAction, int uiParam, string path, int fWinIni);
        protected void ChangeBackground(int idWallpaper)
        {
            SystemParametersInfo(0x0014, 0, this.pathWallpapers + "bg" + idWallpaper.ToString() + ".png", 0x0001);
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
            PowerShell ps = PowerShell.Create();

            ps.AddScript(PATH_PS + "vd_check.ps1");
            var results = ps.Invoke();

            foreach (var psObject in results)
            {
                string buf = psObject.ToString();

                if (buf.All(char.IsDigit))
                {
                    int nBuf = Int16.Parse(buf);

                    if (this.acceptWorkspace.Contains(nBuf))
                    {
                        this.idCurrentWorkspace = nBuf;
                        SetTrayIconN(nBuf);
                    }
                }
            }
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
