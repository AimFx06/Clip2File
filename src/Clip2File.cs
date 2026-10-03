// Clip2File - 把剪贴板里的截图保存到你当前打开的那个文件夹
// v0.1.1 | MIT License
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
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace Clip2File
{
    internal static class AppInfo
    {
        public const string Name = "Clip2File";
        public const string Version = "0.1.1";
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

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();

        public const int WM_HOTKEY = 0x0312;
        public const uint GA_ROOT = 2;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;
        public const uint MOD_SHIFT = 0x0004;
        public const uint MOD_NOREPEAT = 0x4000;
        public static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);

        /// <summary>窗口是不是「资源管理器文件夹窗口」。</summary>
        public static bool IsExplorerWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return false;
            try
            {
                StringBuilder sb = new StringBuilder(256);
                int n = GetClassName(hwnd, sb, sb.Capacity);
                if (n <= 0) return false;
                string cls = sb.ToString();
                return cls == "CabinetWClass" || cls == "ExploreWClass";
            }
            catch { return false; }
        }
    }

    internal static class Settings
    {
        private static readonly string DirPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppInfo.DirName);

        private static string ReadFile(string name)
        {
            try
            {
                string p = Path.Combine(DirPath, name);
                if (File.Exists(p)) return File.ReadAllText(p, Encoding.UTF8).Trim();
            }
            catch { }
            return string.Empty;
        }

        private static void WriteFile(string name, string value)
        {
            try
            {
                Directory.CreateDirectory(DirPath);
                File.WriteAllText(Path.Combine(DirPath, name), value ?? string.Empty, Encoding.UTF8);
            }
            catch { }
        }

        public static string LastDir
        {
            get { return ReadFile("last.txt"); }
            set { WriteFile("last.txt", value); }
        }

        /// <summary>识别不到打开着的文件夹时，图片落到这里。默认：图片\Clip2File</summary>
        public static string DefaultDir
        {
            get
            {
                string d = ReadFile("default.txt");
                if (!string.IsNullOrEmpty(d)) return d;
                return BuiltinDefaultDir();
            }
            set { WriteFile("default.txt", value); }
        }

        public static string BuiltinDefaultDir()
        {
            try
            {
                string pics = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                if (!string.IsNullOrEmpty(pics)) return Path.Combine(pics, AppInfo.Name);
            }
            catch { }
            try
            {
                string desk = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (!string.IsNullOrEmpty(desk)) return Path.Combine(desk, AppInfo.Name);
            }
            catch { }
            return string.Empty;
        }
    }

    /// <summary>
    /// 读取资源管理器文件夹窗口的路径。
    /// 注意：Shell.Application 是 COM 对象（System.__ComObject），
    /// 用 Type.InvokeMember 反射取属性会失败，必须走 IDispatch 后期绑定，
    /// 这里用 Microsoft.VisualBasic.Interaction.CallByName 实现。
    /// </summary>
    internal static class ShellWindows
    {
        private static List<KeyValuePair<IntPtr, string>> Enumerate()
        {
            List<KeyValuePair<IntPtr, string>> list = new List<KeyValuePair<IntPtr, string>>();
            object shell = null;
            try
            {
                shell = Interaction.CreateObject("Shell.Application", string.Empty);
                object windows = Interaction.CallByName(shell, "Windows", CallType.Get, null);
                IEnumerable en = windows as IEnumerable;
                if (en == null) return list;

                foreach (object w in en)
                {
                    try
                    {
                        object h = Interaction.CallByName(w, "HWND", CallType.Get, null);
                        if (h == null) continue;
                        IntPtr hwnd = new IntPtr(Convert.ToInt64(h));
                        if (hwnd == IntPtr.Zero) continue;

                        string p = null;
                        try
                        {
                            object doc = Interaction.CallByName(w, "Document", CallType.Get, null);
                            object folder = Interaction.CallByName(doc, "Folder", CallType.Get, null);
                            object self = Interaction.CallByName(folder, "Self", CallType.Get, null);
                            p = Interaction.CallByName(self, "Path", CallType.Get, null) as string;
                        }
                        catch { }

                        if (!string.IsNullOrEmpty(p) && Directory.Exists(p))
                            list.Add(new KeyValuePair<IntPtr, string>(hwnd, p));
                    }
                    catch { }
                }
            }
            catch { }
            finally
            {
                try { if (shell != null) Marshal.ReleaseComObject(shell); }
                catch { }
            }
            return list;
        }

        public static string PathOfWindow(IntPtr root)
        {
            if (root == IntPtr.Zero) return null;
            foreach (KeyValuePair<IntPtr, string> kv in Enumerate())
                if (kv.Key == root) return kv.Value;
            return null;
        }

        /// <summary>前台窗口若是文件夹窗口，返回它的路径。</summary>
        public static string ForegroundPath()
        {
            IntPtr fg = Native.GetForegroundWindow();
            if (!Native.IsExplorerWindow(fg)) return null;
            return PathOfWindow(Native.GetAncestor(fg, Native.GA_ROOT));
        }

        /// <summary>系统里只开了一个文件夹窗口时，直接用它。</summary>
        public static string SingleWindowPath()
        {
            List<KeyValuePair<IntPtr, string>> list = Enumerate();
            return list.Count == 1 ? list[0].Value : null;
        }
    }

    internal static class Saver
    {
        public static bool EnsureDir(string d)
        {
            if (string.IsNullOrEmpty(d)) return false;
            try { if (!Directory.Exists(d)) Directory.CreateDirectory(d); }
            catch { }
            return Directory.Exists(d);
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
        private readonly Timer watcher = new Timer();
        private readonly Label lblState = new Label();
        private readonly Label lblHint = new Label();
        private readonly Label lblRule = new Label();
        private readonly Label lblDefCaption = new Label();
        private readonly Label lblLastCaption = new Label();
        private readonly TextBox txtDef = new TextBox();
        private readonly TextBox txtLast = new TextBox();
        private readonly Label lblFooter = new Label();
        private readonly Button btnSave = new Button();
        private readonly Button btnPick = new Button();
        private readonly Button btnDef = new Button();
        private readonly CheckBox chkTray = new CheckBox();

        private string hotkeyText = "（未注册）";
        private bool hotkeyOk;
        private bool reallyExit;
        private IntPtr lastExplorerHwnd = IntPtr.Zero;
        private IntPtr lastForeground = IntPtr.Zero;
        private string lastKnownDir;

        public MainForm()
        {
            BuildUi();
            BuildTray();
            RegisterHotkeys();
            RefreshLabels();
            StartWatcher();
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
            ClientSize = new Size(440, 336);
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { }

            lblState.AutoSize = true;
            lblState.Font = new Font(Font.FontFamily, 12F, FontStyle.Bold);
            lblState.ForeColor = Color.FromArgb(32, 150, 72);
            lblState.Location = new Point(20, 16);
            lblState.Text = "● 正在运行";

            lblHint.AutoSize = true;
            lblHint.ForeColor = Color.FromArgb(80, 80, 80);
            lblHint.Location = new Point(22, 48);
            lblHint.Text = string.Empty;

            lblRule.AutoSize = true;
            lblRule.ForeColor = Color.FromArgb(140, 140, 140);
            lblRule.Location = new Point(22, 70);
            lblRule.Text = "图片存进当前打开的文件夹，没有则存下面";

            lblDefCaption.AutoSize = true;
            lblDefCaption.ForeColor = Color.FromArgb(120, 120, 120);
            lblDefCaption.Location = new Point(22, 104);
            lblDefCaption.Text = "默认位置";

            txtDef.ReadOnly = true;
            txtDef.BorderStyle = BorderStyle.FixedSingle;
            txtDef.BackColor = Color.FromArgb(248, 249, 251);
            txtDef.ForeColor = Color.FromArgb(60, 60, 60);
            txtDef.Location = new Point(20, 122);
            txtDef.Size = new Size(320, 25);
            txtDef.TabStop = false;

            btnDef.Text = "更改…";
            btnDef.Location = new Point(348, 121);
            btnDef.Size = new Size(74, 27);
            btnDef.FlatStyle = FlatStyle.System;
            btnDef.Click += delegate { ChangeDefaultDir(); };

            lblLastCaption.AutoSize = true;
            lblLastCaption.ForeColor = Color.FromArgb(120, 120, 120);
            lblLastCaption.Location = new Point(22, 162);
            lblLastCaption.Text = "上次保存到";

            txtLast.ReadOnly = true;
            txtLast.BorderStyle = BorderStyle.FixedSingle;
            txtLast.BackColor = Color.FromArgb(248, 249, 251);
            txtLast.ForeColor = Color.FromArgb(60, 60, 60);
            txtLast.Location = new Point(20, 180);
            txtLast.Size = new Size(402, 25);
            txtLast.TabStop = false;

            btnSave.Text = "立即保存";
            btnSave.Location = new Point(20, 216);
            btnSave.Size = new Size(120, 34);
            btnSave.FlatStyle = FlatStyle.System;
            btnSave.Click += delegate { DoSave(false); };

            btnPick.Text = "另选文件夹…";
            btnPick.Location = new Point(150, 216);
            btnPick.Size = new Size(130, 34);
            btnPick.FlatStyle = FlatStyle.System;
            btnPick.Click += delegate { DoSave(true); };

            chkTray.AutoSize = true;
            chkTray.Location = new Point(20, 262);
            chkTray.Checked = true;
            chkTray.Text = "关闭窗口时留在托盘继续工作";

            lblFooter.AutoSize = true;
            lblFooter.ForeColor = Color.FromArgb(155, 155, 155);
            lblFooter.Location = new Point(20, 296);
            lblFooter.Text = "托盘图标右键可退出　·　" + AppInfo.Name + " v" + AppInfo.Version;

            Controls.Add(lblState);
            Controls.Add(lblHint);
            Controls.Add(lblRule);
            Controls.Add(lblDefCaption);
            Controls.Add(txtDef);
            Controls.Add(btnDef);
            Controls.Add(lblLastCaption);
            Controls.Add(txtLast);
            Controls.Add(btnSave);
            Controls.Add(btnPick);
            Controls.Add(chkTray);
            Controls.Add(lblFooter);
        }

        private void BuildTray()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("保存剪贴板图片", null, delegate { DoSave(false); });
            menu.Items.Add("另选文件夹并保存…", null, delegate { DoSave(true); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("打开主窗口", null, delegate { ShowMain(); });
            menu.Items.Add("打开默认位置", null, delegate { OpenDir(Settings.DefaultDir); });
            menu.Items.Add("打开上次保存的位置", null, delegate { OpenDir(Settings.LastDir); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("退出", null, delegate { ExitApp(); });

            tray.ContextMenuStrip = menu;
            tray.Visible = true;
            try { tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { tray.Icon = SystemIcons.Application; }
            tray.Text = AppInfo.Name;
            tray.DoubleClick += delegate { ShowMain(); };
        }

        // -------------------------------------------------------- watcher
        /// <summary>
        /// 记住"最近一个打开过的文件夹窗口"。这样即使按热键时前台是本程序窗口
        /// 或截图工具，也仍然知道该存到哪里。只在窗口句柄变化时才查一次。
        /// </summary>
        private void StartWatcher()
        {
            watcher.Interval = 500;
            watcher.Tick += delegate
            {
                try
                {
                    IntPtr fg = Native.GetForegroundWindow();
                    if (fg == lastForeground) return;
                    lastForeground = fg;
                    if (!Native.IsExplorerWindow(fg)) return;

                    lastExplorerHwnd = Native.GetAncestor(fg, Native.GA_ROOT);
                    string p = ShellWindows.PathOfWindow(lastExplorerHwnd);
                    if (!string.IsNullOrEmpty(p)) lastKnownDir = p;
                }
                catch { }
            };
            watcher.Start();
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
                ? "截图后按 " + hotkeyText + " 保存"
                : "热键被占用，请用「另选文件夹…」手动保存";
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
        /// <summary>决定存到哪里：当前文件夹 → 最近打开过的文件夹 → 唯一的文件夹窗口 → 默认位置</summary>
        private string ResolveTargetDir(out bool usedDefault)
        {
            usedDefault = false;

            string dir = ShellWindows.ForegroundPath();
            if (!string.IsNullOrEmpty(dir)) return dir;

            if (lastExplorerHwnd != IntPtr.Zero)
            {
                dir = ShellWindows.PathOfWindow(lastExplorerHwnd);
                if (!string.IsNullOrEmpty(dir)) return dir;
            }

            if (!string.IsNullOrEmpty(lastKnownDir) && Directory.Exists(lastKnownDir)) return lastKnownDir;

            dir = ShellWindows.SingleWindowPath();
            if (!string.IsNullOrEmpty(dir)) return dir;

            usedDefault = true;
            return Settings.DefaultDir;
        }

        private void DoSave(bool pickFolder)
        {
            string dir;
            bool usedDefault = false;
            string tail = string.Empty;

            if (pickFolder)
            {
                if (!Visible) ShowMain();
                dir = PickFolder();
                if (string.IsNullOrEmpty(dir)) return;
            }
            else
            {
                dir = ResolveTargetDir(out usedDefault);
                if (!Saver.EnsureDir(dir))
                {
                    Balloon("没有可用的保存位置", "请在主窗口里设置一个默认位置", ToolTipIcon.Warning);
                    if (!Visible) ShowMain();
                    return;
                }
                if (usedDefault) tail = "\n（默认位置）";
            }

            byte[] bytes = Saver.ClipboardImageBytes();
            if (bytes != null && bytes.Length > 0)
            {
                try
                {
                    string path = Saver.SaveImage(dir, bytes);
                    Settings.LastDir = dir;
                    RefreshLabels();
                    Balloon("已保存截图", Path.GetFileName(path) + "\n" + dir + tail, ToolTipIcon.Info);
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
                RefreshLabels();
                Balloon("已复制文件", copied.Count + " 个文件 → " + dir + tail, ToolTipIcon.Info);
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

        private void ChangeDefaultDir()
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = "选择识别不到文件夹时要用的默认保存位置";
                dlg.ShowNewFolderButton = true;
                string cur = Settings.DefaultDir;
                if (!string.IsNullOrEmpty(cur) && Directory.Exists(cur)) dlg.SelectedPath = cur;
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    Settings.DefaultDir = dlg.SelectedPath;
                    RefreshLabels();
                }
            }
        }

        private void OpenDir(string dir)
        {
            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
            {
                try { System.Diagnostics.Process.Start("explorer.exe", "\"" + dir + "\""); }
                catch { }
            }
            else Balloon("目录不存在", "还没有保存过，或目录已被删除", ToolTipIcon.Info);
        }

        private void RefreshLabels()
        {
            string def = Settings.DefaultDir;
            txtDef.Text = string.IsNullOrEmpty(def) ? "（未设置）" : def;

            string last = Settings.LastDir;
            txtLast.Text = string.IsNullOrEmpty(last) ? "（还没有保存过）" : last;
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
                try { watcher.Stop(); watcher.Dispose(); } catch { }
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
                    string dir = args.Length > 1 ? args[1] : ShellWindows.ForegroundPath();
                    if (string.IsNullOrEmpty(dir)) dir = Settings.DefaultDir;
                    if (!Saver.EnsureDir(dir)) return;

                    byte[] bytes = Saver.ClipboardImageBytes();
                    if (bytes != null && bytes.Length > 0)
                    {
                        Saver.SaveImage(dir, bytes);
                        Settings.LastDir = dir;
                        return;
                    }

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
