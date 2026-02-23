using System.Collections.Generic;

namespace AutoClickScenarioTool.Models
{
    public class ScenarioStep
    {
        public int Delay { get; set; }
        public List<string> Positions { get; set; } = new List<string>();
    }
}
