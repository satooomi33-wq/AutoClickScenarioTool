using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;
using AutoClickScenarioTool.Models;

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
                SendKey(vk, useScanCode, false);
                Thread.Sleep(10);
                SendKey(vk, useScanCode, true);

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

            SendInput(1, inputs, INPUT.Size);
        }

        // Simple helper to click a sequence of points (synchronous)
        public void ClickMultiple(PositionList list)
        {
            if (list == null) return;
            try
            {
                foreach (var p in list.Points)
                {
                    SetCursorPos(p.X, p.Y);
                    // press
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(10);
                    // release
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(30);
                }
            }
            catch { }
        }

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    }
}
