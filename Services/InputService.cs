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
    }

    public class PositionList
    {
        public System.Collections.Generic.List<Point> Points { get; } = new System.Collections.Generic.List<Point>();
    }
}
