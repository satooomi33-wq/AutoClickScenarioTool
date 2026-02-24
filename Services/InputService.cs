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

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint KEYEVENTF_KEYUP = 0x0002;

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
                byte vk = ConvertKeyNameToVk(main);
                if (vk != 0)
                {
                    keybd_event(vk, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(20);
                    keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
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
