using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace AutoClickScenarioTool.Services
{
    public class GlobalKeyboardHook : IDisposable
    {
        public event Action<string>? OnKeyPressed;

        private IntPtr _hookId = IntPtr.Zero;
        private HookProc _proc;

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
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule!)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            const int WM_KEYDOWN = 0x0100;
            const int WM_SYSKEYDOWN = 0x0104;
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                string keyName = ((System.Windows.Forms.Keys)vkCode).ToString();
                // normalize common names
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
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private const int WH_KEYBOARD_LL = 13;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public void Dispose() => Stop();
    }
}
