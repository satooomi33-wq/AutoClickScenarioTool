using System;

namespace AutoClickScenarioTool.Models
{
    public class DefaultSettings
    {
        public int Delay { get; set; } = 500;
        public int PressDuration { get; set; } = 100;
        // 擬人化（ヒューマナイズ）設定
        public bool HumanizeEnabled { get; set; } = true;
        public int HumanizeLower { get; set; } = 30;
        public int HumanizeUpper { get; set; } = 100;

        // New: persist user's preferred runtime SC mode
        public bool UseScanCode { get; set; } = false;
    }
}
