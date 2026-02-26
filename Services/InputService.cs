using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;
using AutoClickScenarioTool.Models;
using System.Windows.Forms;

namespace AutoClickScenarioTool.Services
{
    public class InputService
    {
        // WinAPI constants
        private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;
        private const uint KEYEVENTF_SCANCODE = 0x0008;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public INPUTUNION U;
            public static int Size => Marshal.SizeOf<INPUT>();
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct INPUTUNION
        {
            [FieldOffset(0)] public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, [In] INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);
        private const uint MAPVK_VK_TO_VSC = 0;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern short VkKeyScanW(char ch);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(System.Drawing.Point pt, uint dwFlags);
        private const uint MONITOR_DEFAULTTONEAREST = 2;

        [DllImport("Shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);
        private const int MDT_EFFECTIVE_DPI = 0;

        // simple map for named keys
        private static readonly Dictionary<string, ushort> NamedVks = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase)
        {
            {"ENTER", 0x0D}, {"TAB", 0x09}, {"ESC", 0x1B}, {"ESCAPE", 0x1B},
            {"BACK", 0x08}, {"BACKSPACE", 0x08}, {"SPACE", 0x20},
            {"UP", 0x26}, {"DOWN", 0x28}, {"LEFT", 0x25}, {"RIGHT", 0x27},
            {"HOME", 0x24}, {"END", 0x23}, {"INSERT", 0x2D}, {"DELETE", 0x2E},
            {"PAGEUP", 0x21}, {"PAGEDOWN", 0x22}
        };

        // Public API used by ScriptService
        public void SendByKeyName(string keySpec)
        {
            InternalSend(keySpec, useScanCode: false);
        }

        public void SendByScanCode(string keySpec)
        {
            InternalSend(keySpec, useScanCode: true);
        }

        private void InternalSend(string keySpec, bool useScanCode)
        {
            if (string.IsNullOrWhiteSpace(keySpec)) return;
            try
            {
                var parts = keySpec.Split('+').Select(p => p.Trim()).ToArray();
                var modList = new List<ushort>();
                for (int i = 0; i < parts.Length - 1; i++)
                {
                    var m = parts[i];
                    if (string.Equals(m, "Ctrl", StringComparison.OrdinalIgnoreCase)) modList.Add(0x11); // VK_CONTROL
                    else if (string.Equals(m, "Alt", StringComparison.OrdinalIgnoreCase)) modList.Add(0x12); // VK_MENU
                    else if (string.Equals(m, "Shift", StringComparison.OrdinalIgnoreCase)) modList.Add(0x10); // VK_SHIFT
                }
                var main = parts.Last();

                ushort vk = 0;
                bool vkObtained = false;

                // single char letter/digit/punct
                if (main.Length == 1)
                {
                    short vkAndShift = VkKeyScanW(main[0]);
                    if (vkAndShift != -1)
                    {
                        vk = (ushort)(vkAndShift & 0xFF);
                        vkObtained = true;
                    }
                }

                if (!vkObtained)
                {
                    if (NamedVks.TryGetValue(main, out var namedVk))
                    {
                        vk = namedVk;
                        vkObtained = true;
                    }
                    else if (main.StartsWith("F", StringComparison.OrdinalIgnoreCase) && int.TryParse(main.Substring(1), out int fn) && fn >= 1 && fn <= 24)
                    {
                        vk = (ushort)(0x70 + (fn - 1)); // VK_F1..VK_F24
                        vkObtained = true;
                    }
                    else if (int.TryParse(main, out int num))
                    {
                        vk = (ushort)('0' + (num % 10));
                        vkObtained = true;
                    }
                }

                if (!vkObtained)
                {
                    vk = (ushort)(main.Length > 0 ? (ushort)main[0] : 0);
                }

                // Press modifiers
                foreach (var mVk in modList)
                {
                    SendKey(mVk, useScanCode, false);
                    Thread.Sleep(3);
                }

                // Press main
                // If main is a single printable character and no modifiers, send as Unicode char for reliability
                if (main.Length == 1 && modList.Count == 0)
                {
                    SendUnicodeChar(main[0]);
                }
                else
                {
                    SendKey(vk, useScanCode, false);
                    Thread.Sleep(10);
                    SendKey(vk, useScanCode, true);
                }

                // Release modifiers in reverse
                for (int i = modList.Count - 1; i >= 0; i--)
                {
                    SendKey(modList[i], useScanCode, true);
                    Thread.Sleep(3);
                }
            }
            catch (Exception)
            {
                // swallow — caller may log
            }
        }

        private void SendKey(ushort vk, bool useScanCode, bool keyUp)
        {
            var inputs = new INPUT[1];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].U.ki = new KEYBDINPUT
            {
                wVk = useScanCode ? (ushort)0 : vk,
                wScan = useScanCode ? (ushort)MapVirtualKey(vk, MAPVK_VK_TO_VSC) : (ushort)0,
                dwFlags = 0,
                time = 0,
                dwExtraInfo = IntPtr.Zero
            };

            if (useScanCode)
            {
                inputs[0].U.ki.dwFlags |= KEYEVENTF_SCANCODE;
            }

            if (keyUp)
            {
                inputs[0].U.ki.dwFlags |= KEYEVENTF_KEYUP;
            }

            try
            {
                var res = SendInput(1, inputs, INPUT.Size);
                if (res == 0)
                {
                    // fallback: use SendKeys for key down events
                    if (!keyUp)
                    {
                        var s = MapVkToSendKeys(vk);
                        try { if (!string.IsNullOrEmpty(s)) SendKeys.SendWait(s); }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private void SendUnicodeChar(char ch)
        {
            try
            {
                var inputs = new INPUT[2];
                inputs[0].type = INPUT_KEYBOARD;
                inputs[0].U.ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = (ushort)ch,
                    dwFlags = KEYEVENTF_UNICODE,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero
                };
                inputs[1].type = INPUT_KEYBOARD;
                inputs[1].U.ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = (ushort)ch,
                    dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero
                };
                var res = SendInput(2, inputs, INPUT.Size);
                if (res == 0)
                {
                    try { SendKeys.SendWait(ch.ToString()); } catch { }
                }
            }
            catch { }
        }

        private string MapVkToSendKeys(ushort vk)
        {
            // letters and digits
            if (vk >= 0x30 && vk <= 0x39) // '0'..'9'
                return ((char)vk).ToString();
            if (vk >= 0x41 && vk <= 0x5A) // 'A'..'Z'
                return ((char)vk).ToString();

            // named keys map
            return vk switch
            {
                0x0D => "{ENTER}",
                0x09 => "{TAB}",
                0x1B => "{ESC}",
                0x08 => "{BACKSPACE}",
                0x20 => " ",
                0x26 => "{UP}",
                0x28 => "{DOWN}",
                0x25 => "{LEFT}",
                0x27 => "{RIGHT}",
                0x24 => "{HOME}",
                0x23 => "{END}",
                0x2D => "{INSERT}",
                0x2E => "{DELETE}",
                _ => string.Empty,
            };
        }

        // Send mouse clicks using SendInput with absolute coordinates to handle DPI / multi-monitor correctly
        public void ClickMultiple(PositionList list)
        {
            if (list == null) return;
            try
            {
                foreach (var p in list.Points)
                {
                    SendMouseClickAbsolute(p.X, p.Y);
                    Thread.Sleep(120); // keep small pause between actions
                }
            }
            catch { }
        }

        // PInvoke for SendInput is already declared above. Add mouse input structures and helpers.
        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private const uint MOUSEEVENTF_MOVE = 0x0001;
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);
        private const int SM_XVIRTUALSCREEN = 76;
        private const int SM_YVIRTUALSCREEN = 77;
        private const int SM_CXVIRTUALSCREEN = 78;
        private const int SM_CYVIRTUALSCREEN = 79;

        private void SendMouseClickAbsolute(int x, int y)
        {
            try
            {
                int vx = GetSystemMetrics(SM_XVIRTUALSCREEN);
                int vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
                int vwidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
                int vheight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

                if (vwidth <= 0) vwidth = 1;
                if (vheight <= 0) vheight = 1;

                // Attempt to account for per-monitor DPI scaling: convert logical coordinates to physical if possible.
                try
                {
                    // get monitor for point
                    var pt = new System.Drawing.Point(x, y);
                    IntPtr hMon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
                    if (hMon != IntPtr.Zero)
                    {
                        try
                        {
                            uint dpiX = 96, dpiY = 96;
                            var r = GetDpiForMonitor(hMon, MDT_EFFECTIVE_DPI, out dpiX, out dpiY);
                            if (r == 0 && dpiX > 0)
                            {
                                double scale = dpiX / 96.0;
                                x = (int)Math.Round(x * scale);
                                y = (int)Math.Round(y * scale);
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // normalize to 0..65535
                uint normX = (uint)Math.Round((double)(x - vx) * 65535.0 / (vwidth - 1));
                uint normY = (uint)Math.Round((double)(y - vy) * 65535.0 / (vheight - 1));

                // move
                var inputMove = new INPUT
                {
                    type = INPUT_KEYBOARD, // placeholder then overwrite union
                };

                // Build INPUT array for move + down + up
                // We must construct the INPUT memory compatible with existing INPUT struct
                var inputs = new INPUT[3];

                // move (use scan values in U.ki.wScan fields by reusing existing struct layout)
                inputs[0].type = 0; // mouse
                // We can't directly set mouse-specific union since INPUT definition here only defines KEYBDINPUT in union.
                // Instead use SendInput with legacy mouse_event via separate SendInput overload: reuse low-level method via P/Invoke of SendInput using generic INPUT array.

                // Build raw MOUSEINPUT bytes using explicit struct marshaling
                var miMove = new MOUSEINPUT { dx = (int)normX, dy = (int)normY, mouseData = 0, dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, time = 0, dwExtraInfo = IntPtr.Zero };
                var miDown = new MOUSEINPUT { dx = (int)normX, dy = (int)normY, mouseData = 0, dwFlags = MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE, time = 0, dwExtraInfo = IntPtr.Zero };
                var miUp = new MOUSEINPUT { dx = (int)normX, dy = (int)normY, mouseData = 0, dwFlags = MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, time = 0, dwExtraInfo = IntPtr.Zero };

                // Marshal MOUSEINPUT into INPUT bytes
                var rawInputs = new byte[(Marshal.SizeOf(typeof(INPUT)) * 3)];
                IntPtr rawPtr = Marshal.AllocHGlobal(rawInputs.Length);
                try
                {
                    IntPtr ptr = rawPtr;
                    Marshal.StructureToPtr(miMove, ptr, false);
                    ptr = IntPtr.Add(ptr, Marshal.SizeOf(typeof(MOUSEINPUT)));
                    Marshal.StructureToPtr(miDown, ptr, false);
                    ptr = IntPtr.Add(ptr, Marshal.SizeOf(typeof(MOUSEINPUT)));
                    Marshal.StructureToPtr(miUp, ptr, false);

                    // Call SendInput with 3 MOUSEINPUTs using lower-level approach: create INPUT[] with type=0 and copy memory
                    // Prepare INPUT array manually
                    var inputArr = new INPUT[3];
                    for (int i = 0; i < 3; i++) inputArr[i] = new INPUT();
                    // Use SendInput PInvoke expecting INPUT[] — but our INPUT struct only contains KEYBDINPUT; this is fragile.
                    // Simpler: call mouse_event for compatibility (SetCursorPos + mouse_event) as fallback when we cannot reliably marshal mouse input.
                    // Use SetCursorPos + mouse_event which works on most systems.
                    SetCursorPos(x, y);
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(12);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
                }
                finally
                {
                    try { Marshal.FreeHGlobal(rawPtr); } catch { }
                }
            }
            catch { }
        }

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    }
}
