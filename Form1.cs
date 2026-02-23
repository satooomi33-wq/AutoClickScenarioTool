using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutoClickScenarioTool.Models;
using AutoClickScenarioTool.Services;

namespace AutoClickScenarioTool
{
    public partial class Form1 : Form
    {
        private readonly DataService _dataService = new DataService();
        private readonly InputService _inputService = new InputService();
        private ScriptService _scriptService;

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
            _scriptService = script;
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

        private void RefreshNoColumn()
        {
            for (int i = 0; i < dgvScenario.Rows.Count; i++)
            {
                var row = dgvScenario.Rows[i];
                if (row.IsNewRow) continue;
                row.Cells[0].Value = (i + 1).ToString();
            }
        }

        private void AppendLog(string s)
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
                        cmbFiles.Items.Add(f);
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
                cells[0] = null; // NO: will be filled
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

                await _scriptService.StartAsync(steps, startIndex).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppendLog("Start error: " + ex.Message);
            }
        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            if (!_scriptService.IsRunning) return;
            if (_isPaused)
            {
                _scriptService.Resume();
                _isPaused = false;
                btnPause.Text = "一時停止";
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
            btnPause.Text = "再開";
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (!_scriptService.IsRunning) return;
            _scriptService.Stop();
            btnStart.Enabled = true;
            btnPause.Enabled = false;
            btnStop.Enabled = false;
            _isPaused = false;
            btnPause.Text = "一時停止";
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
            btnPause.Enabled = false;
            btnStop.Enabled = false;
            _isPaused = false;
            btnPause.Text = "一時停止";
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

        private async void btnCapture_Click(object sender, EventArgs e)
        {
            AppendLog("座標抽出中: 画面上で左クリックしてください...");
            await Task.Run(async () =>
            {
                // wait for left button down
                while ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0)
                {
                    await Task.Delay(30).ConfigureAwait(false);
                }
                var p = System.Windows.Forms.Cursor.Position;
                Invoke(new Action(() =>
                {
                    int rowIndex = dgvScenario.CurrentCell?.RowIndex ?? -1;
                    if (rowIndex < 0 || rowIndex >= dgvScenario.Rows.Count)
                    {
                        dgvScenario.Rows.Add();
                        rowIndex = dgvScenario.Rows.Count - 1;
                    }

                    // find first pos column that is empty or use current cell if in pos column
                    int targetCol = -1;
                    if (dgvScenario.CurrentCell != null && dgvScenario.CurrentCell.ColumnIndex >= 2)
                        targetCol = dgvScenario.CurrentCell.ColumnIndex;
                    else
                    {
                        for (int c = 2; c < dgvScenario.ColumnCount; c++)
                        {
                            var v = dgvScenario.Rows[rowIndex].Cells[c].Value?.ToString();
                            if (string.IsNullOrWhiteSpace(v)) { targetCol = c; break; }
                        }
                    }

                    if (targetCol == -1)
                    {
                        MessageBox.Show("空きの座標列がありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // ウィンドウを前面に戻す
                    SetForegroundWindow(this.Handle);
                    // セルに座標を書き込む
                    dgvScenario.Rows[rowIndex].Cells[targetCol].Value = $"{p.X},{p.Y}";
                    RefreshNoColumn();
                    AppendLog($"座標取得: {p.X},{p.Y}");
                    // フォーカスを遅延で強制復帰
                    BeginInvoke(new Action(() => {
                        dgvScenario.CurrentCell = dgvScenario.Rows[rowIndex].Cells[targetCol];
                        dgvScenario.Select();
                        dgvScenario.Focus();
                    }));
                }));
            }).ConfigureAwait(false);
        }
    }
}
