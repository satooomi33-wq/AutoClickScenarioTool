using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;

namespace AutoClickScenarioTool.Services
{
    public class GlobalKeyboardHook : IDisposable
    {
        public event Action<string>? OnKeyPressed;

        // true の場合、処理済みのキー入力を他のアプリやコントロールへ渡さない（抑止）
        public bool SuppressKeys { get; set; } = false;
        // 実行時に現在キャプチャしたキーを抑止するか決めるオプションの述語
        public Func<bool>? ShouldSuppress { get; set; }

        private IntPtr _hookId = IntPtr.Zero;
        private HookProc _proc;

        // コンストラクタ
        public GlobalKeyboardHook()
        {
            _proc = HookCallback;
        }

        public void Start()
        {
            _hookId = SetHook(_proc);
        }

        public void Stop()
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        private IntPtr SetHook(HookProc proc)
        {
            // For low-level keyboard hooks it's more reliable to pass IntPtr.Zero
            // as the module handle when installing from managed code.
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc, IntPtr.Zero, 0);
        }

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        // フックコールバック: キーイベントを受け取って処理・変換し、必要なら suppression を行う
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                const int WM_KEYDOWN = 0x0100;
                const int WM_SYSKEYDOWN = 0x0104;
                if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    var key = (System.Windows.Forms.Keys)vkCode;
                    // 純粋な修飾キー（Shift, Ctrl, Alt）だけの押下は無視する
                    if (key == Keys.ShiftKey || key == Keys.ControlKey || key == Keys.Menu || key == Keys.LShiftKey || key == Keys.RShiftKey || key == Keys.LControlKey || key == Keys.RControlKey || key == Keys.LMenu || key == Keys.RMenu)
                    {
                        return CallNextHookEx(_hookId, nCode, wParam, lParam);
                    }

                    string keyName = key.ToString();
                    // Oem系キー（例: OemPlus）の場合、可能なら実際の文字に変換を試みる
                    if (keyName.StartsWith("Oem", StringComparison.OrdinalIgnoreCase))
                    {
                        if (TryGetCharFromKey((int)key, out string ch) && !string.IsNullOrEmpty(ch))
                        {
                            keyName = ch;
                        }
                        else
                        {
                            // fallback map for common OEM VKs
                            keyName = MapOemKeyName((int)key) ?? keyName;
                        }
                    }
                    // 一般的なキー名を正規化
                    if (keyName.Length == 2 && keyName[0] == 'D' && char.IsDigit(keyName[1]))
                    {
                        keyName = keyName.Substring(1);
                    }
                    else if (keyName.StartsWith("NumPad", StringComparison.OrdinalIgnoreCase))
                    {
                        keyName = keyName.Substring(6);
                    }
                    else if (string.Equals(keyName, "Return", StringComparison.OrdinalIgnoreCase))
                    {
                        keyName = "Enter";
                    }

                    OnKeyPressed?.Invoke(keyName);

                    // 抑止は、SuppressKeys=true のときにのみ行う。かつ ShouldSuppress が true を返すか
                    // またはこのプロセスが前景フォアグラウンドのときに抑止する。
                    if (SuppressKeys && OnKeyPressed != null)
                    {
                        try
                        {
                            // Avoid suppressing when Windows key is held (prevents Win+Shift+S etc)
                            const int VK_LWIN = 0x5B;
                            const int VK_RWIN = 0x5C;
                            if ((GetAsyncKeyState(VK_LWIN) & 0x8000) != 0 || (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0)
                            {
                                // let it through
                            }
                            else
                            {
                                // If a consumer provided a ShouldSuppress predicate, respect it.
                                if (ShouldSuppress != null)
                                {
                                    try
                                    {
                                        if (ShouldSuppress())
                                            return (IntPtr)1;
                                    }
                                    catch { }
                                }
                                else
                                {
                                    var fg = GetForegroundWindow();
                                    if (fg != IntPtr.Zero)
                                    {
                                        GetWindowThreadProcessId(fg, out uint pid);
                                        if (pid == (uint)Process.GetCurrentProcess().Id)
                                        {
                                            return (IntPtr)1; // suppress further processing
                                        }
                                    }
                                }
                            }
                        }
                        catch { /* swallow to avoid blocking */ }
                    }
                }
            }
            catch
            {
                // never let hook throw — always pass through
                try { return CallNextHookEx(_hookId, nCode, wParam, lParam); } catch { return IntPtr.Zero; }
            }

            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private static string? MapOemKeyName(int vk)
        {
            return vk switch
            {
                0xBB => "+", // VK_OEM_PLUS
                0xBC => ",", // VK_OEM_COMMA
                0xBD => "-", // VK_OEM_MINUS
                0xBE => ".", // VK_OEM_PERIOD
                0xBF => "/", // VK_OEM_2
                0xBA => ";", // VK_OEM_1
                0xDE => "'", // VK_OEM_7
                0xC0 => "`", // VK_OEM_3
                0xE2 => "\\", // VK_OEM_102
                0xDC => "\\", // VK_OEM_5
                _ => null,
            };
        }

        private static bool TryGetCharFromKey(int vk, out string result)
        {
            result = string.Empty;
            try
            {
                var ks = new byte[256];
                if (!GetKeyboardState(ks)) return false;
                uint scan = MapVirtualKey((uint)vk, 0);
                var sb = new StringBuilder(8);
                int rc = ToUnicode((uint)vk, scan, ks, sb, sb.Capacity, 0);
                if (rc > 0)
                {
                    result = sb.ToString();
                    return true;
                }
            }
            catch { }
            return false;
        }

        private const int WH_KEYBOARD_LL = 13;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState, [Out, MarshalAs(UnmanagedType.LPWStr, SizeParamIndex = 4)] StringBuilder pwszBuff, int cchBuff, uint wFlags);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public void Dispose() => Stop();
    }
}
