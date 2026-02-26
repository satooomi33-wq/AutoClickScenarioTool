using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace AutoClickScenarioTool.Services
{
    public class KeyboardMonitor : IDisposable
    {
        private delegate IntPtr LowLevelProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelProc? _proc;
        private IntPtr _hook = IntPtr.Zero;

        public event Action<int, int, int>? OnKey; // vkCode, scanCode, flags

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr GetModuleHandle(string? lpModuleName);

        public void Start()
        {
            if (_hook != IntPtr.Zero) return;
            _proc = HookCallback;
            // get module handle for current process
            IntPtr mod = GetModuleHandle(Process.GetCurrentProcess().MainModule?.ModuleName);
            _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, mod, 0);
        }

        public void Stop()
        {
            if (_hook == IntPtr.Zero) return;
            try { UnhookWindowsHookEx(_hook); } catch { }
            _hook = IntPtr.Zero;
            _proc = null;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0)
                {
                    int wm = wParam.ToInt32();
                    if (wm == WM_KEYDOWN || wm == WM_KEYUP)
                    {
                        var ks = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                        OnKey?.Invoke((int)ks.vkCode, (int)ks.scanCode, (int)ks.flags);
                    }
                }
            }
            catch { }
            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            Stop();
        }
    }
}
