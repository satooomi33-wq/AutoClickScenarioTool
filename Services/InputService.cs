using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;

namespace AutoClickScenarioTool.Services
{
    public class InputService
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public INPUTUNION u;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct INPUTUNION
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, [In] INPUT[] pInputs, int cbSize);

        private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_UNICODE = 0x0004;
        
        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_SCANCODE = 0x0008;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

        public void ClickAt(Point p)
        {
            System.Windows.Forms.Cursor.Position = p;
            Thread.Sleep(5);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(5);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        public void ClickMultiple(PositionList positions)
        {
            // Helper: click each position quickly to simulate simultaneous presses.
            foreach (var p in positions.Points)
            {
                ClickAt(p);
                Thread.Sleep(8);
            }
        }

        // キー送信: use keybd_event to synthesize key presses (more reliable than SendKeys)
        public void SendKey(string keySpec)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(keySpec)) return;
                var parts = keySpec.Split('+').Select(p => p.Trim()).Where(p => p.Length > 0).ToArray();
                var mods = new System.Collections.Generic.List<System.Windows.Forms.Keys>();
                string main = parts.Last();
                for (int i = 0; i < parts.Length - 1; i++)
                {
                    var m = parts[i];
                    if (string.Equals(m, "Ctrl", StringComparison.OrdinalIgnoreCase) || string.Equals(m, "Control", StringComparison.OrdinalIgnoreCase)) mods.Add(System.Windows.Forms.Keys.ControlKey);
                    else if (string.Equals(m, "Alt", StringComparison.OrdinalIgnoreCase)) mods.Add(System.Windows.Forms.Keys.Menu);
                    else if (string.Equals(m, "Shift", StringComparison.OrdinalIgnoreCase)) mods.Add(System.Windows.Forms.Keys.ShiftKey);
                }

                // Press modifiers
                foreach (var mk in mods)
                {
                    keybd_event((byte)mk, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(10);
                }

                // Main key
                // Support explicit scan-code mode via prefix "SC:" (e.g., "SC:A" or "SC:65").
                if (main.StartsWith("SC:", StringComparison.OrdinalIgnoreCase))
                {
                    var rem = main.Substring(3).Trim();
                    uint scan = 0;
                    if (!uint.TryParse(rem, out scan))
                    {
                        // try as key name -> VK -> scan
                        byte vk2 = ConvertKeyNameToVk(rem);
                        if (vk2 != 0)
                        {
                            scan = MapVirtualKey(vk2, 0);
                        }
                    }

                    if (scan != 0)
                    {
                        var inputs = new System.Collections.Generic.List<INPUT>();
                        var down = new INPUT
                        {
                            type = INPUT_KEYBOARD,
                            u = new INPUTUNION { ki = new KEYBDINPUT { wVk = 0, wScan = (ushort)scan, dwFlags = KEYEVENTF_SCANCODE, time = 0, dwExtraInfo = UIntPtr.Zero } }
                        };
                        var up = new INPUT
                        {
                            type = INPUT_KEYBOARD,
                            u = new INPUTUNION { ki = new KEYBDINPUT { wVk = 0, wScan = (ushort)scan, dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP, time = 0, dwExtraInfo = UIntPtr.Zero } }
                        };
                        inputs.Add(down);
                        inputs.Add(up);
                        SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(INPUT)));
                    }
                }
                else
                {
                    byte vk = ConvertKeyNameToVk(main);
                    if (vk != 0)
                    {
                        // Use keybd_event for known virtual-key codes (shortcuts, control keys, function keys)
                        keybd_event(vk, 0, 0, UIntPtr.Zero);
                        Thread.Sleep(20);
                        keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                    }
                    else
                    {
                        // Fallback: send as Unicode characters (handles IME and non-ASCII input)
                        var inputs = new System.Collections.Generic.List<INPUT>();
                        foreach (char ch in main)
                        {
                            var kiDown = new INPUT
                            {
                                type = INPUT_KEYBOARD,
                                u = new INPUTUNION { ki = new KEYBDINPUT { wVk = 0, wScan = (ushort)ch, dwFlags = KEYEVENTF_UNICODE, time = 0, dwExtraInfo = UIntPtr.Zero } }
                            };
                            var kiUp = new INPUT
                            {
                                type = INPUT_KEYBOARD,
                                u = new INPUTUNION { ki = new KEYBDINPUT { wVk = 0, wScan = (ushort)ch, dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, time = 0, dwExtraInfo = UIntPtr.Zero } }
                            };
                            inputs.Add(kiDown);
                            inputs.Add(kiUp);
                        }
                        if (inputs.Count > 0)
                        {
                            SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(INPUT)));
                        }
                    }
                }

                // Release modifiers
                for (int i = mods.Count - 1; i >= 0; i--)
                {
                    keybd_event((byte)mods[i], 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                    Thread.Sleep(10);
                }
            }
            catch { }
        }

        private byte ConvertKeyNameToVk(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return 0;
            var up = name.Trim().ToUpperInvariant();
            // single char A-Z or 0-9
            if (up.Length == 1 && char.IsLetterOrDigit(up[0])) return (byte)up[0];
            switch (up)
            {
                case "ENTER": return 0x0D;
                case "TAB": return 0x09;
                case "ESC":
                case "ESCAPE": return 0x1B;
                case "BACK":
                case "BACKSPACE": return 0x08;
                case "SPACE": return 0x20;
                case "LEFT": return 0x25;
                case "UP": return 0x26;
                case "RIGHT": return 0x27;
                case "DOWN": return 0x28;
                default:
                    if (up.StartsWith("F") && int.TryParse(up.Substring(1), out int fn))
                    {
                        if (fn >= 1 && fn <= 24) return (byte)(0x70 + (fn - 1));
                    }
                    if (int.TryParse(up, out int num))
                    {
                        // digits
                        if (num >= 0 && num <= 9) return (byte)('0' + num);
                    }
                    break;
            }
            return 0;
        }
    }

    public class PositionList
    {
        public System.Collections.Generic.List<Point> Points { get; } = new System.Collections.Generic.List<Point>();
    }
}
