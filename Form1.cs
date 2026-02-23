using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using AutoClickScenarioTool.Models;
using AutoClickScenarioTool.Services;

namespace AutoClickScenarioTool
{
    public partial class Form1 : Form
    {
        private readonly DataService _dataService = new DataService();
        private readonly InputService _inputService = new InputService();
        // ...existing code...
        private ScriptService? _scriptService;

        private bool _isPaused = false;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01;

        public Form1()
        {
            InitializeComponent();

            _script_service_init();

            CreateGridColumns();

            // プロジェクト直下のDataを優先し、なければbin配下を使う
            try
            {
                var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                var projDir = Path.GetFullPath(Path.Combine(exeDir, "..", "..", ".."));
                var projData = Path.Combine(projDir, "Data");
                if (Directory.Exists(projData) && Directory.EnumerateFiles(projData, "*.json").Any())
                {
                    txtDataFolder.Text = projData;
                }
                else
                {
                    var binData = Path.Combine(exeDir, "Data");
                    if (!Directory.Exists(binData))
                        Directory.CreateDirectory(binData);
                    txtDataFolder.Text = binData;
                }
            }
            catch { /* ignore */ }

            RefreshFileList();
        }

        private void _script_service_init()
        {
            _script_service_init_inner();
        }

        private void _script_service_init_inner()
        {
            // initialize script service and hookup
            var script = new ScriptService(_input_service_wrap());
            _scriptService = script!;
            _scriptService.OnLog += s => AppendLog(s);
            _scriptService.OnStopped += () => Invoke(new Action(ScriptStopped));
            _scriptService.OnPaused += (idx) => Invoke(new Action<int>(HandlePaused));
        }

        // wrappers to avoid tiny naming collisions
        private InputService _input_service_wrap() => _input_service_real();
        private InputService _input_service_real() => _inputService;

        private void CreateGridColumns()
        {
            dgvScenario.Columns.Clear();
            var colNo = new DataGridViewTextBoxColumn { Name = "NO", HeaderText = "NO", ReadOnly = true, Width = 40 };
            var colDelay = new DataGridViewTextBoxColumn { Name = "Delay", HeaderText = "遅延(ms)", Width = 80 };
            dgvScenario.Columns.Add(colNo);
            dgvScenario.Columns.Add(colDelay);
            for (int i = 1; i <= 10; i++)
            {
                dgvScenario.Columns.Add(new DataGridViewTextBoxColumn { Name = $"Pos{i}", HeaderText = $"座標{i}", Width = 120 });
            }
        }

        public void RefreshNoColumn()
        {
            for (int i = 0; i < dgvScenario.Rows.Count; i++)
            {
                var row = dgvScenario.Rows[i];
                if (row.IsNewRow) continue;
                row.Cells[0].Value = (i + 1).ToString();
            }
        }

        public void AppendLog(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendLog), s);
                return;
            }
            txtLog.AppendText(s + Environment.NewLine);
        }

        private void RefreshFileList()
        {
            try
            {
                var folder = txtDataFolder.Text;
                cmbFiles.Items.Clear();
                cmbFiles.Text = ""; // 初期値ブランク
                if (Directory.Exists(folder))
                {
                    foreach (var f in _dataService.ListJsonFiles(folder))
                    {
                        var name = Path.GetFileNameWithoutExtension(f);
                        cmbFiles.Items.Add(name);
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("Error listing files: " + ex.Message);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtDataFolder.Text = dlg.SelectedPath;
                RefreshFileList();
            }
        }

        private void cmbFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadScenarioFromCombo();
        }

        private void cmbFiles_TextChanged(object sender, EventArgs e)
        {
            LoadScenarioFromCombo();
        }

        private async void LoadScenarioFromCombo()
        {
            var folder = txtDataFolder.Text;
            var fileName = cmbFiles.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(folder) || string.IsNullOrWhiteSpace(fileName)) return;
            if (!fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                fileName += ".json";
            var path = Path.Combine(folder, fileName);
            if (!File.Exists(path)) return;
            try
            {
                var steps = await _data_service_load(path).ConfigureAwait(false);
                Invoke(new Action(() => FillGridWithSteps(steps)));
                AppendLog("読み込み完了: " + path);
            }
            catch (Exception ex)
            {
                AppendLog("読み込みエラー: " + ex.Message);
            }
        }

        private Task<List<ScenarioStep>> _data_service_load(string path) => _dataService.LoadAsync(path);

        private async void btnSave_Click(object sender, EventArgs e)
        {
            var folder = txtDataFolder.Text;
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("DATAフォルダを指定してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string? fileName = cmbFiles.Text?.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                MessageBox.Show("ファイル名を入力してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                fileName += ".json";
            string path = Path.Combine(folder, fileName);

            var steps = ReadStepsFromGrid();
            bool fileExists = File.Exists(path);
            DialogResult dr;
            if (fileExists)
            {
                dr = MessageBox.Show($"{fileName} は既存ファイルです。上書きしますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;
            }
            else
            {
                dr = MessageBox.Show($"{fileName} を新規作成します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;
            }
            try
            {
                await _dataService.SaveAsync(path, steps).ConfigureAwait(false);
                AppendLog("保存: " + path);
                RefreshFileList();
            }
            catch (Exception ex)
            {
                AppendLog("保存エラー: " + ex.Message);
            }
        }

        private void FillGridWithSteps(List<ScenarioStep> steps)
        {
            dgvScenario.Rows.Clear();
            foreach (var s in steps)
            {
                var cells = new object[12];
                cells[0] = string.Empty; // NO: will be filled
                cells[1] = s.Delay;
                for (int i = 0; i < 10; i++)
                {
                    cells[2 + i] = i < s.Positions.Count ? s.Positions[i] : string.Empty;
                }
                dgvScenario.Rows.Add(cells);
            }
            RefreshNoColumn();
        }


        private List<ScenarioStep> ReadStepsFromGrid()
        {
            var list = new List<ScenarioStep>();
            foreach (DataGridViewRow row in dgvScenario.Rows)
            {
                if (row.IsNewRow) continue;
                var step = new ScenarioStep();
                var delayCell = row.Cells[1].Value;
                if (delayCell != null && int.TryParse(delayCell.ToString(), out var d))
                    step.Delay = d;
                else
                    step.Delay = 0;

                for (int i = 0; i < 10; i++)
                {
                    var v = row.Cells[2 + i].Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(v))
                        step.Positions.Add(v.Trim());
                }
                list.Add(step);
            }
            return list;
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                var steps = ReadStepsFromGrid();
                if (steps.Count == 0)
                {
                    AppendLog("シナリオが空です。");
                    return;
                }

                int startIndex = 0;
                if (dgvScenario.CurrentCell != null)
                    startIndex = dgvScenario.CurrentCell.RowIndex;

                AppendLog($"実行開始: 行{startIndex + 1} から");
                _isPaused = false;
                btnStart.Enabled = false;
                btnPause.Enabled = true;
                btnStop.Enabled = true;

                await _scriptService!.StartAsync(steps, startIndex).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppendLog("Start error: " + ex.Message);
            }
        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            if (!(_scriptService?.IsRunning ?? false)) return;
            if (_isPaused)
            {
                _scriptService.Resume();
                _isPaused = false;
            }
            else
            {
                _script_service_pause();
            }
        }

        private void _script_service_pause()
        {
            _scriptService.Pause();
            _isPaused = true;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (!(_scriptService?.IsRunning ?? false)) return;
            _scriptService.Stop();
            btnStart.Enabled = true;
            btnPause.Enabled = false;
            btnStop.Enabled = false;
            _isPaused = false;
            if (dgvScenario.Rows.Count > 0)
            {
                _script_service_pause();
                // bring app front when user requests pause
                try { SetForegroundWindow(this.Handle); } catch { }
            }
        }

        private void ScriptStopped()
        {
            // bring app front
            try { SetForegroundWindow(this.Handle); } catch { }
            btnStart.Enabled = true;
            btnPause.Enabled = true;
            btnStop.Enabled = true;
            _isPaused = false;
            // focus first row if present
            if (dgvScenario.Rows.Count > 0)
            {
                try { dgvScenario.CurrentCell = dgvScenario.Rows[0].Cells[1]; dgvScenario.Focus(); } catch { }
            }
            AppendLog("スクリプト停止");
        }

        private void HandlePaused(int idx)
        {
            try { SetForegroundWindow(this.Handle); } catch { }
            if (idx >= 0 && idx < dgvScenario.Rows.Count)
            {
                try
                {
                    dgvScenario.CurrentCell = dgvScenario.Rows[idx].Cells[1];
                    dgvScenario.Focus();
                }
                catch { }
            }
        }

        private void dgvScenario_RowsChanged(object sender, EventArgs e)
        {
            RefreshNoColumn();
        }

        // セル編集後、最下行にデータが入ったら新たな空行を追加
        private void dgvScenario_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // NO列を常に更新
            RefreshNoColumn();

            // 編集された行が最下行（新規行の直前）かつ、何か値が入った場合
            if (e.RowIndex == dgvScenario.Rows.Count - 2 && !dgvScenario.Rows[e.RowIndex].IsNewRow)
            {
                bool hasData = false;
                for (int c = 1; c < dgvScenario.ColumnCount; c++)
                {
                    var v = dgvScenario.Rows[e.RowIndex].Cells[c].Value;
                    if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) { hasData = true; break; }
                }
                if (hasData && dgvScenario.AllowUserToAddRows && dgvScenario.Rows[dgvScenario.Rows.Count - 1].IsNewRow)
                {
                    // 編集確定を促すことで新規行が追加される
                    dgvScenario.EndEdit();
                }
            }

        }

        // ユーザーが新規行を追加したときもNO列を更新
        private void dgvScenario_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            RefreshNoColumn();
        }

        private bool _captureMode = false;
        private IntPtr _mainHandle;
        private GlobalMouseHook? _mouseHook;
        // 重複定義削除
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            _mainHandle = this.Handle;
            btnToggleCapture.CheckedChanged += BtnToggleCapture_CheckedChanged;
            // _mouseClickFilter の初期化は不要（フィールド未定義のため）
            _mouseHook = new GlobalMouseHook();
            _mouseHook.OnMouseClick += HandleGlobalMouseClick;
        }

        private void BtnToggleCapture_CheckedChanged(object? sender, EventArgs e)
        {
            _captureMode = btnToggleCapture.Checked;
            btnToggleCapture.Text = _captureMode ? "座標抽出 ON" : "座標抽出 OFF";
            if (_captureMode)
            {
                _mouseHook?.Start();
                AppendLog("座標抽出モード: ON");
            }
            else
            {
                _mouseHook?.Stop();
                AppendLog("座標抽出モード: OFF");
            }

        }

        // グローバルマウスクリック時の座標抽出処理
        public void HandleGlobalMouseClick(int x, int y)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => HandleGlobalMouseClick(x, y)));
                return;
            }
            // 座標カラムのセルにフォーカスがある場合のみ抽出
            if (dgvScenario.CurrentCell == null || dgvScenario.CurrentCell.ColumnIndex < 2) return;
            int rowIndex = dgvScenario.CurrentCell.RowIndex;
            int targetCol = dgvScenario.CurrentCell.ColumnIndex;
            if (dgvScenario.IsCurrentCellInEditMode)
                dgvScenario.EndEdit();
            dgvScenario.Rows[rowIndex].Cells[targetCol].Value = $"{x},{y}";
            RefreshNoColumn();
            AppendLog($"座標取得: {x},{y}");
            if (rowIndex == dgvScenario.Rows.Count - 2 && dgvScenario.AllowUserToAddRows)
            {
                dgvScenario.Rows.Add();
                RefreshNoColumn();
            }
            BeginInvoke(new Action(() => {
                dgvScenario.CurrentCell = dgvScenario.Rows[rowIndex].Cells[targetCol];
                dgvScenario.Select();
                dgvScenario.Focus();
            }));
        }

    // グローバルマウスクリック監視用MessageFilter
public class MouseClickMessageFilter : IMessageFilter
{
    private readonly Form1 _form;
    public MouseClickMessageFilter(Form1 form) { _form = form; }
    public bool PreFilterMessage(ref Message m)
    {
        const int WM_LBUTTONDOWN = 0x0201;
        if (m.Msg == WM_LBUTTONDOWN)
        {
            var cursorPos = System.Windows.Forms.Cursor.Position;
            // 条件を緩和：どこをクリックしても抽出する
            if (_form.dgvScenario.CurrentCell == null) return false;
            int rowIndex = _form.dgvScenario.CurrentCell.RowIndex;
            int targetCol = _form.dgvScenario.CurrentCell.ColumnIndex;
            if (_form.dgvScenario.IsCurrentCellInEditMode)
                _form.dgvScenario.EndEdit();
            _form.dgvScenario.Rows[rowIndex].Cells[targetCol].Value = $"{cursorPos.X},{cursorPos.Y}";
            _form.RefreshNoColumn();
            _form.AppendLog($"座標取得: {cursorPos.X},{cursorPos.Y}");
            if (rowIndex == _form.dgvScenario.Rows.Count - 2 && _form.dgvScenario.AllowUserToAddRows)
            {
                _form.dgvScenario.Rows.Add();
                _form.RefreshNoColumn();
            }
            _form.BeginInvoke(new Action(() => {
                _form.dgvScenario.CurrentCell = _form.dgvScenario.Rows[rowIndex].Cells[targetCol];
                _form.dgvScenario.Select();
                _form.dgvScenario.Focus();
            }));
        }
        return false;
    }

}
// Form1 クラス閉じ括弧
}
// クラス閉じ括弧追加

}
