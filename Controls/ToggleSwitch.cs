using System.Windows.Forms;

namespace AutoClickScenarioTool.Controls
{
    public class ToggleSwitch : CheckBox
    {
        public ToggleSwitch()
        {
            this.Appearance = Appearance.Button;
            this.Text = "OFF";
            this.AutoSize = false;
            this.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Width = 110;
            this.Height = 45;
            this.CheckedChanged += (s, e) =>
            {
                this.Text = this.Checked ? "ON" : "OFF";
            };
        }
    }
}
