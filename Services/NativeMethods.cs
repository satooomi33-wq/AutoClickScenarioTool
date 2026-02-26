using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace AutoClickScenarioTool.Services
{
    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern IntPtr MonitorFromPoint(System.Drawing.Point pt, uint dwFlags);
        internal const uint MONITOR_DEFAULTTONEAREST = 2;

        [DllImport("Shcore.dll")]
        internal static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);
        internal const int MDT_EFFECTIVE_DPI = 0;
    }
}
