// Clip2File - 把剪贴板里的截图保存到你当前打开的那个文件夹
// v0.1.0 | MIT License
//
// 为什么需要它：Windows 截图后图片只进剪贴板（位图格式），而文件夹的"粘贴"
// 只认文件列表格式，所以粘不进去。本工具补上这一步。
//
// 用法：双击运行 → 托盘出现图标 → 截图后按 Ctrl+Alt+V → 图片落到当前文件夹。

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Clip2File
{
    internal static class AppInfo
    {
        public const string Name = "Clip2File";
        public const string Version = "0.1.0";
        public const string MutexName = "Clip2File_SingleInstance_v1";
        public const string DirName = "Clip2File";
    }

    internal static class Native
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();

        public const int WM_HOTKEY = 0x0312;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;
        public const uint MOD_SHIFT = 0x0004;
        public const uint MOD_NOREPEAT = 0x4000;
        public static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);
    }

    internal static class Settings
    {
        private static readonly string DirPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppInfo.DirName);
        private static readonly string FilePath = Path.Combine(DirPath, "settings.ini");

        public static string LastDir
        {
            get
            {
                try
                {
                    if (File.Exists(FilePath)) return File.ReadAllText(FilePath, Encoding.UTF8).Trim();
                }
                catch { }
                return string.Empty;
            }
            set
            {
                try
                {
                    Directory.CreateDirectory(DirPath);
                    File.WriteAllText(FilePath, value ?? string.Empty, Encoding.UTF8);
                }
                catch { }
            }
        }
    }

    internal static class Saver
    {
        /// <summary>取当前前台资源管理器窗口所在的文件夹；不是文件夹窗口则返回 null。</summary>
        public static string GetForegroundFolder()
        {
            try
            {
                IntPtr hwnd = Native.GetForegroundWindow();
                if (hwnd == IntPtr.Zero) return null;

                Type shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType == null) return null;

                object shell = Activator.CreateInstance(shellType);
                try
                {
                    IEnumerable windows = Prop(shell, "Windows") as IEnumerable;
                    if (windows == null) return null;

                    foreach (object w in windows)
                    {
                        try
                        {
                            IntPtr wh = new IntPtr(Convert.ToInt64(Prop(w, "HWND")));
                            if (wh != hwnd) continue;

                            object doc = Prop(w, "Document");
                            object folder = Prop(doc, "Folder");
                            object self = Prop(folder, "Self");
                            string p = Prop(self, "Path") as string;
                            if (!string.IsNullOrEmpty(p) && Directory.Exists(p)) return p;
                        }
                        catch { }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(shell);
                }
            }
            catch { }
            return null;
        }

        private static object Prop(object o, string name)
        {
            return o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, null);
        }

        /// <summary>把剪贴板里的图片取成 PNG 字节；没有图片返回 null。</summary>
        public static byte[] ClipboardImageBytes()
        {
            try
            {
                IDataObject data = Clipboard.GetDataObject();
                if (data != null && data.GetDataPresent("PNG"))
                {
                    object o = data.GetData("PNG", true);
                    Stream s = o as Stream;
                    if (s != null)
                    {
                        s.Position = 0;
                        using (MemoryStream ms = new MemoryStream())
                        {
                            s.CopyTo(ms);
                            return ms.ToArray();
                        }
                    }
                    byte[] b = o as byte[];
                    if (b != null) return b;
                }
            }
            catch { }

            try
            {
                using (Image img = Clipboard.GetImage())
                {
                    if (img != null)
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            img.Save(ms, ImageFormat.Png);
                            return ms.ToArray();
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        public static string SaveImage(string dir, byte[] bytes)
        {
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string path = Path.Combine(dir, "截图_" + stamp + ".png");
            int i = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, "截图_" + stamp + "_" + i + ".png");
                i++;
            }
            File.WriteAllBytes(path, bytes);
            return path;
        }

        /// <summary>剪贴板里是"复制的文件"时，把它们复制到目标文件夹。</summary>
        public static List<string> CopyClipboardFiles(string dir)
        {
            List<string> saved = new List<string>();
            StringCollection list = null;
            try { list = Clipboard.GetFileDropList(); }
            catch { }
            if (list == null) return saved;

            foreach (string f in list)
            {
                try
                {
                    if (string.IsNullOrEmpty(f) || !File.Exists(f)) continue;

                    string dest = Path.Combine(dir, Path.GetFileName(f));
                    if (File.Exists(dest))
                    {
                        string bn = Path.GetFileNameWithoutExtension(f);
                        string ex = Path.GetExtension(f);
                        int i = 1;
                        do
                        {
                            dest = Path.Combine(dir, bn + "_(" + i + ")" + ex);
                            i++;
                        } while (File.Exists(dest));
                    }
                    File.Copy(f, dest, true);
                    saved.Add(dest);
                }
                catch { }
            }
            return saved;
        }
    }

    internal sealed class MainForm : Form
    {
        private const int IdSave = 1;
        private const int IdShow = 2;

        private static readonly uint MsgShow = Native.RegisterWindowMessage("Clip2File.ShowMainWindow");

        private readonly NotifyIcon tray = new NotifyIcon();
        private readonly Label lblState = new Label();
        private readonly Label lblHint = new Label();
        private readonly Label lblLastCaption = new Label();
        private readonly Label lblLast = new Label();
        private readonly Label lblFooter = new Label();
        private readonly Button btnSave = new Button();
        private readonly Button btnPick = new Button();
        private readonly CheckBox chkTray = new CheckBox();

        private string hotkeyText = "（未注册）";
        private bool hotkeyOk;
        private bool reallyExit;

        public MainForm()
        {
            BuildUi();
            BuildTray();
            RegisterHotkeys();
            RefreshLastLabel();
            Shown += delegate { WarnIfHotkeyFailed(); };
        }

        // ------------------------------------------------------------ UI
        private void BuildUi()
        {
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = AppInfo.Name + " " + AppInfo.Version;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(430, 268);
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { }

            lblState.AutoSize = true;
            lblState.Font = new Font(Font.FontFamily, 12F, FontStyle.Bold);
            lblState.ForeColor = Color.FromArgb(32, 150, 72);
            lblState.Location = new Point(20, 18);
            lblState.Text = "● 正在运行";

            lblHint.AutoSize = true;
            lblHint.ForeColor = Color.FromArgb(90, 90, 90);
            lblHint.Location = new Point(22, 50);
            lblHint.Text = string.Empty;

            lblLastCaption.AutoSize = true;
            lblLastCaption.ForeColor = Color.FromArgb(120, 120, 120);
            lblLastCaption.Location = new Point(22, 108);
            lblLastCaption.Text = "上次保存到";

            lblLast.AutoSize = false;
            lblLast.BorderStyle = BorderStyle.FixedSingle;
            lblLast.BackColor = Color.FromArgb(248, 249, 251);
            lblLast.ForeColor = Color.FromArgb(60, 60, 60);
            lblLast.Location = new Point(20, 128);
            lblLast.Size = new Size(390, 30);
            lblLast.TextAlign = ContentAlignment.MiddleLeft;
            lblLast.Padding = new Padding(8, 0, 0, 0);
            lblLast.AutoEllipsis = true;

            btnSave.Text = "立即保存";
            btnSave.Location = new Point(20, 174);
            btnSave.Size = new Size(120, 34);
            btnSave.FlatStyle = FlatStyle.System;
            btnSave.Click += delegate { DoSave(false); };

            btnPick.Text = "选择文件夹…";
            btnPick.Location = new Point(150, 174);
            btnPick.Size = new Size(120, 34);
            btnPick.FlatStyle = FlatStyle.System;
            btnPick.Click += delegate { DoSave(true); };

            chkTray.AutoSize = true;
            chkTray.Location = new Point(288, 182);
            chkTray.Checked = true;
            chkTray.Text = "关闭窗口时留在托盘";

            lblFooter.AutoSize = true;
            lblFooter.ForeColor = Color.FromArgb(150, 150, 150);
            lblFooter.Location = new Point(20, 226);
            lblFooter.Text = "托盘图标右键可退出　·　" + AppInfo.Name + " v" + AppInfo.Version;

            Controls.Add(lblState);
            Controls.Add(lblHint);
            Controls.Add(lblLastCaption);
            Controls.Add(lblLast);
            Controls.Add(btnSave);
            Controls.Add(btnPick);
            Controls.Add(chkTray);
            Controls.Add(lblFooter);
        }

        private void BuildTray()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("保存剪贴板图片", null, delegate { DoSave(false); });
            menu.Items.Add("选择文件夹并保存…", null, delegate { DoSave(true); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("打开主窗口", null, delegate { ShowMain(); });
            menu.Items.Add("打开上次保存的位置", null, delegate { OpenLast(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("退出", null, delegate { ExitApp(); });

            tray.ContextMenuStrip = menu;
            tray.Visible = true;
            try { tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { tray.Icon = SystemIcons.Application; }
            tray.Text = AppInfo.Name;
            tray.DoubleClick += delegate { ShowMain(); };
        }

        // ------------------------------------------------------- hotkeys
        private void RegisterHotkeys()
        {
            hotkeyOk = Native.RegisterHotKey(Handle, IdSave, Native.MOD_CONTROL | Native.MOD_ALT | Native.MOD_NOREPEAT, 0x56);
            if (hotkeyOk) hotkeyText = "Ctrl + Alt + V";
            else
            {
                hotkeyOk = Native.RegisterHotKey(Handle, IdSave, Native.MOD_CONTROL | Native.MOD_ALT | Native.MOD_NOREPEAT, 0x58);
                if (hotkeyOk) hotkeyText = "Ctrl + Alt + X";
                else
                {
                    hotkeyOk = Native.RegisterHotKey(Handle, IdSave, Native.MOD_CONTROL | Native.MOD_ALT | Native.MOD_NOREPEAT, 0x5A);
                    if (hotkeyOk) hotkeyText = "Ctrl + Alt + Z";
                }
            }

            Native.RegisterHotKey(Handle, IdShow, Native.MOD_CONTROL | Native.MOD_ALT | Native.MOD_SHIFT | Native.MOD_NOREPEAT, 0x51);

            lblHint.Text = hotkeyOk
                ? "截图后按 " + hotkeyText + " ，图片直接存进你当前打开的文件夹"
                : "热键被其他软件占用了，请点「选择文件夹…」手动保存";
            tray.Text = AppInfo.Name + "　" + hotkeyText;
        }

        private void WarnIfHotkeyFailed()
        {
            if (hotkeyOk) return;
            MessageBox.Show(this,
                "无法注册全局热键（Ctrl+Alt+V / X / Z 都被占用）。\n\n" +
                "你仍然可以用窗口里的按钮或托盘菜单保存，\n" +
                "关掉占用热键的软件后重启本程序即可恢复。",
                AppInfo.Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Native.WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                if (id == IdSave) DoSave(false);
                else if (id == IdShow) ShowMain();
            }
            else if (MsgShow != 0 && (uint)m.Msg == MsgShow)
            {
                ShowMain();
            }
            base.WndProc(ref m);
        }

        // -------------------------------------------------------- actions
        private void DoSave(bool forceDialog)
        {
            string dir = forceDialog ? null : Saver.GetForegroundFolder();

            if (string.IsNullOrEmpty(dir))
            {
                if (!Visible) ShowMain();
                dir = PickFolder();
            }
            if (string.IsNullOrEmpty(dir)) return;

            byte[] bytes = Saver.ClipboardImageBytes();
            if (bytes != null && bytes.Length > 0)
            {
                try
                {
                    string path = Saver.SaveImage(dir, bytes);
                    Settings.LastDir = dir;
                    RefreshLastLabel();
                    Balloon("已保存截图", Path.GetFileName(path) + "\n" + dir, ToolTipIcon.Info);
                    return;
                }
                catch (Exception ex)
                {
                    Balloon("保存失败", ex.Message, ToolTipIcon.Error);
                    return;
                }
            }

            List<string> copied = Saver.CopyClipboardFiles(dir);
            if (copied.Count > 0)
            {
                Settings.LastDir = dir;
                RefreshLastLabel();
                Balloon("已复制文件", copied.Count + " 个文件 → " + dir, ToolTipIcon.Info);
                return;
            }

            Balloon("剪贴板里没有图片", "先截图（Win+Shift+S）或复制文件，再按 " + hotkeyText, ToolTipIcon.Warning);
        }

        private string PickFolder()
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = "选择要保存到的文件夹";
                dlg.ShowNewFolderButton = true;
                string last = Settings.LastDir;
                if (!string.IsNullOrEmpty(last) && Directory.Exists(last)) dlg.SelectedPath = last;
                return dlg.ShowDialog(this) == DialogResult.OK ? dlg.SelectedPath : null;
            }
        }

        private void OpenLast()
        {
            string last = Settings.LastDir;
            if (!string.IsNullOrEmpty(last) && Directory.Exists(last))
            {
                try { System.Diagnostics.Process.Start("explorer.exe", "\"" + last + "\""); }
                catch { }
            }
            else Balloon("还没有保存记录", "先保存一张截图试试", ToolTipIcon.Info);
        }

        private void RefreshLastLabel()
        {
            string last = Settings.LastDir;
            lblLast.Text = string.IsNullOrEmpty(last) ? "（还没有保存过）" : last;
        }

        private void Balloon(string title, string text, ToolTipIcon icon)
        {
            try { tray.ShowBalloonTip(3500, title, text, icon); }
            catch { }
        }

        private void ShowMain()
        {
            Show();
            if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
            Activate();
            BringToFront();
        }

        private void ExitApp()
        {
            reallyExit = true;
            Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!reallyExit && e.CloseReason == CloseReason.UserClosing && chkTray.Checked)
            {
                e.Cancel = true;
                Hide();
                Balloon(AppInfo.Name + " 还在运行", "按 " + hotkeyText + " 保存截图；右键托盘图标可退出", ToolTipIcon.Info);
                return;
            }
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { Native.UnregisterHotKey(Handle, IdSave); } catch { }
                try { Native.UnregisterHotKey(Handle, IdShow); } catch { }
                try { tray.Visible = false; tray.Dispose(); } catch { }
            }
            base.Dispose(disposing);
        }
    }

    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Native.SetProcessDPIAware();

            // ---------- 命令行模式：Clip2File.exe --save "D:\某个文件夹" ----------
            if (args.Length > 0)
            {
                string a0 = args[0].ToLowerInvariant();

                if (a0 == "--save" || a0 == "-s")
                {
                    string dir = args.Length > 1 ? args[1] : Saver.GetForegroundFolder();
                    if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return;

                    byte[] bytes = Saver.ClipboardImageBytes();
                    if (bytes != null && bytes.Length > 0) { Saver.SaveImage(dir, bytes); Settings.LastDir = dir; return; }

                    List<string> copied = Saver.CopyClipboardFiles(dir);
                    if (copied.Count > 0) Settings.LastDir = dir;
                    return;
                }

                if (a0 == "--help" || a0 == "-h" || a0 == "/?")
                {
                    MessageBox.Show(
                        AppInfo.Name + " " + AppInfo.Version + "\n\n" +
                        "用法：\n" +
                        "  双击运行            启动托盘程序，截图后按 Ctrl+Alt+V\n" +
                        "  Clip2File.exe --save <文件夹>    静默保存一次后退出\n\n" +
                        "保存规则：\n" +
                        "  · 剪贴板里是图片 → 存成 截图_时间戳.png\n" +
                        "  · 剪贴板里是文件 → 复制这些文件到目标文件夹\n",
                        AppInfo.Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (a0 == "--version" || a0 == "-v")
                {
                    MessageBox.Show(AppInfo.Name + " " + AppInfo.Version, AppInfo.Name,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            // ---------- 单实例 ----------
            bool createdNew;
            using (System.Threading.Mutex mutex = new System.Threading.Mutex(true, AppInfo.MutexName, out createdNew))
            {
                if (!createdNew)
                {
                    uint msg = Native.RegisterWindowMessage("Clip2File.ShowMainWindow");
                    if (msg != 0) Native.PostMessage(Native.HWND_BROADCAST, msg, IntPtr.Zero, IntPtr.Zero);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
                GC.KeepAlive(mutex);
            }
        }
    }
}
