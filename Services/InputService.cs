using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;

namespace AutoClickScenarioTool.Services
{
    public class InputService
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

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

        // キー送信: simple conversion to SendKeys format
        public void SendKey(string keySpec)
        {
            try
            {
                var send = ConvertToSendKeysFormat(keySpec);
                if (!string.IsNullOrEmpty(send))
                    System.Windows.Forms.SendKeys.SendWait(send);
            }
            catch { }
        }

        private string ConvertToSendKeysFormat(string keySpec)
        {
            if (string.IsNullOrWhiteSpace(keySpec)) return string.Empty;
            var parts = keySpec.Split('+');
            var modifiers = new System.Collections.Generic.List<string>();
            string main = parts.Last().Trim();
            for (int i = 0; i < parts.Length - 1; i++)
            {
                var m = parts[i].Trim();
                if (string.Equals(m, "Ctrl", StringComparison.OrdinalIgnoreCase) || string.Equals(m, "Control", StringComparison.OrdinalIgnoreCase)) modifiers.Add("^");
                else if (string.Equals(m, "Alt", StringComparison.OrdinalIgnoreCase)) modifiers.Add("%");
                else if (string.Equals(m, "Shift", StringComparison.OrdinalIgnoreCase)) modifiers.Add("+");
            }

            string mainSend;
            if (main.Length == 1)
            {
                mainSend = main;
            }
            else if (int.TryParse(main, out _))
            {
                mainSend = main;
            }
            else
            {
                switch (main.ToUpperInvariant())
                {
                    case "ENTER": mainSend = "{ENTER}"; break;
                    case "TAB": mainSend = "{TAB}"; break;
                    case "ESC":
                    case "ESCAPE": mainSend = "{ESC}"; break;
                    case "BACK":
                    case "BACKSPACE": mainSend = "{BACKSPACE}"; break;
                    case "SPACE": mainSend = " "; break;
                    default:
                        if (main.StartsWith("F", StringComparison.OrdinalIgnoreCase) && int.TryParse(main.Substring(1), out _))
                            mainSend = "{" + main.ToUpperInvariant() + "}";
                        else
                            mainSend = main;
                        break;
                }
            }

            return string.Join(string.Empty, modifiers) + mainSend;
        }
    }

    public class PositionList
    {
        public System.Collections.Generic.List<Point> Points { get; } = new System.Collections.Generic.List<Point>();
    }
}
