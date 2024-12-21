using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Drawing;
using System.Management;
using Microsoft.Win32;

namespace NoSleep
{
    public partial class NoSleepForm : Form
    {
        // スクリーンセーバー待ち時間取得用
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, out uint pvParam, uint fWinIni);
        private const uint SPI_GETSCREENSAVETIMEOUT = 14;

        // マウスイベント関連の定数
        const int MOUSEEVENTF_MOVED = 0x0001;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;

        // アンマネージ DLL 対応用 struct 記述宣言
        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public int type;
            public MOUSEINPUT mi;
        }

        [DllImport("user32.dll")]
        extern static uint SendInput(
            uint nInputs,
            INPUT[] pInputs,
            int cbSize
        );

        // タイマー設定
        private Timer screensaverTimer;
        private ManagementEventWatcher screensaverWatcher;

        // キーボードフック関連の定義
        private static LowLevelKeyboardProc _keyboardProc = KeyboardHookCallback;
        private static IntPtr _keyboardHookID = IntPtr.Zero;

        // マウスフック関連の定義
        private static LowLevelMouseProc _mouseProc = MouseHookCallback;
        private static IntPtr _mouseHookID = IntPtr.Zero;

        private static NoSleepForm _instance;

        public NoSleepForm()
        {
            InitializeComponent();
            _instance = this; // 静的インスタンスを初期化

            screensaverTimer = new Timer();
            screensaverTimer.Tick += ScreensaverTimer_Tick;

            // 初回の待ち時間設定
            UpdateScreensaverTimer();

            // スクリーンセーバー設定変更を監視する
            StartScreensaverWatcher();

            // グローバルキーボードフックを設定
            _keyboardHookID = SetKeyboardHook(_keyboardProc);

            // グローバルマウスフックを設定
            _mouseHookID = SetMouseHook(_mouseProc);

            Application.ApplicationExit += new EventHandler(OnApplicationExit);
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            UnhookWindowsHookEx(_keyboardHookID);
            UnhookWindowsHookEx(_mouseHookID);
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr SetKeyboardHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr SetMouseHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Debug.WriteLine($"キーボードの入力を検知：{(Keys)vkCode}");

                // キーが押されたときにスクリーンセーバータイマーをリセット
                _instance.ResetScreensaverTimer();
            }
            return CallNextHookEx(_keyboardHookID, nCode, wParam, lParam);
        }

        private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_MOUSEMOVE)
            {
                Debug.WriteLine("マウスの移動を検知");
                // マウスが動いたときにスクリーンセーバータイマーをリセット
                _instance.ResetScreensaverTimer();
            }
            return CallNextHookEx(_mouseHookID, nCode, wParam, lParam);
        }

        public void ResetScreensaverTimer()
        {
            screensaverTimer.Stop();
            screensaverTimer.Start();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private const int WH_KEYBOARD_LL = 13;
        private const int WH_MOUSE_LL = 14;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_MOUSEMOVE = 0x0200;

        private void StartScreensaverWatcher()
        {
            try
            {
                // スクリーンセーバー設定の変更を監視するWMIクエリ
                string query = "SELECT * FROM __InstanceModificationEvent " +
                               "WITHIN 2 WHERE TargetInstance ISA 'Win32_Desktop'";

                screensaverWatcher = new ManagementEventWatcher(new WqlEventQuery(query));
                screensaverWatcher.EventArrived += OnScreensaverSettingsChanged;
                screensaverWatcher.Start();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"スクリーンセーバー監視のセットアップ中にエラーが発生しました: {ex.Message}");
            }
        }

        private void OnScreensaverSettingsChanged(object sender, EventArrivedEventArgs e)
        {
            Debug.WriteLine("スクリーンセーバー設定が変更されました");

            // 設定が変更されたため、タイマーを更新
            UpdateScreensaverTimer();
        }

        private async void ScreensaverTimer_Tick(object sender, EventArgs e)
        {
            await autoMouseRunStart(2000);    // 2秒待つ
        }

        private void UpdateScreensaverTimer()
        {
            int interval = 50 * 1000;   // 初期値(50秒)

            if (IsScreensaverNone())
            {
                Debug.WriteLine("スクリーンセーバーが設定されていない");
            }
            else if (SystemParametersInfo(SPI_GETSCREENSAVETIMEOUT, 0, out uint timeoutSeconds, 0))
            {
                interval = (int)(timeoutSeconds - 10) * 1000;
                Debug.WriteLine($"スクリーンセーバーの待ち時間 - 10秒");
            }

            Debug.WriteLine($"screensaverTimer.Interval：{interval / 1000}秒");
            screensaverTimer.Interval = interval;
            screensaverTimer.Start();
        }

        static bool IsScreensaverNone()
        {
            const string keyName = @"HKEY_CURRENT_USER\Control Panel\Desktop";
            const string valueName = "SCRNSAVE.EXE";

            // レジストリから値を取得
            object scrnsaveValue = Registry.GetValue(keyName, valueName, null);

            // 値が存在しないか空文字列なら「(なし)」
            return scrnsaveValue == null || string.IsNullOrEmpty(scrnsaveValue.ToString());
        }

        private void AppExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            screensaverTimer.Stop();
            appExit();
        }

        private async Task autoMouseRunStart(int duration)
        {
            (int x, int y) = GetAbsoluteMousePosition();

            INPUT[] input = new INPUT[2];
            input[0].mi.dx = x + 100;
            input[0].mi.dy = y;
            input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;

            input[1].mi.dx = x;
            input[1].mi.dy = y;
            input[1].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;

            await Task.Run(async () =>
            {
                var endTime = DateTime.Now.AddMilliseconds(duration);
                while (DateTime.Now < endTime)
                {
                    SendInput(2, input, Marshal.SizeOf(input[0]));
                    await Task.Delay(10);
                }
            });                                                                 
        }

        private (int X, int Y) GetAbsoluteMousePosition()
        {
            Point cursorPosition = Cursor.Position;
            Screen currentScreen = Screen.FromPoint(cursorPosition);
            int x = (cursorPosition.X * (65535 / currentScreen.Bounds.Width));
            int y = (cursorPosition.Y * (65535 / currentScreen.Bounds.Height));

            return (x, y);
        }

        private void appExit()
        {
            notifyIcon.Visible = false;
            screensaverWatcher?.Stop();
            screensaverWatcher?.Dispose();
            Application.Exit();
        }
    }
}
