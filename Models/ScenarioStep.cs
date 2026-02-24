using System.Collections.Generic;

namespace AutoClickScenarioTool.Models
{
    public class ScenarioStep
    {
        public int Delay { get; set; }
        // 押下（クリック/タッチ）時間（ミリ秒）
        public int PressDuration { get; set; } = 100;
        public List<string> Positions { get; set; } = new List<string>();
        // 新: 混在アクション（座標文字列 "X,Y" または キー指定 "A" / "Ctrl+S" など）
        public List<string> Actions { get; set; } = new List<string>();
    }
}
