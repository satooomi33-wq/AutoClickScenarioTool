using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Threading;
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

        

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_SHOWWINDOW = 0x0040;

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(System.Drawing.Point p);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);
        private const uint GA_ROOT = 2;

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(System.Drawing.Point pt, uint dwFlags);
        private const uint MONITOR_DEFAULTTONEAREST = 2;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AllowSetForegroundWindow(uint dwProcessId);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("Shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);
        private const int MDT_EFFECTIVE_DPI = 0;

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private const int SW_SHOWNORMAL = 1;
        private const byte VK_MENU = 0x12; // ALT
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private static readonly IntPtr HWND_TOP = new IntPtr(0);

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

        // Designer-referenced handlers: add simple implementations to avoid CS1061
        private void btnAddRow_Click(object? sender, EventArgs e)
        {
            // insert new row after current
            try
            {
                int idx = dgvScenario.CurrentCell?.RowIndex ?? dgvScenario.Rows.Count - 1;
                InsertBlankRowAt(idx + 1);
            }
            catch { }
        }

        private void btnRowUp_Click(object? sender, EventArgs e)
        {
            try
            {
                if (dgvScenario.CurrentCell == null) return;
                var r = dgvScenario.CurrentCell.RowIndex;
                MoveRow(r, Math.Max(0, r - 1));
            }
            catch { }
        }

        private void btnRowDown_Click(object? sender, EventArgs e)
        {
            try
            {
                if (dgvScenario.CurrentCell == null) return;
                var r = dgvScenario.CurrentCell.RowIndex;
                MoveRow(r, Math.Min(dgvScenario.Rows.Count - 1, r + 1));
            }
            catch { }
        }

        private void LogDisplayInfo()
        {
            try
            {
                AppendLog("--- Display Info Start ---");
                foreach (var s in Screen.AllScreens)
                {
                    AppendLog($"Screen: {s.DeviceName}, Primary={s.Primary}, Bounds={s.Bounds}, WorkingArea={s.WorkingArea}");
                    var pt = new System.Drawing.Point(s.Bounds.Left + 10, s.Bounds.Top + 10);
                    var hMon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
                    if (hMon != IntPtr.Zero)
                    {
                        try
                        {
                            var res = GetDpiForMonitor(hMon, MDT_EFFECTIVE_DPI, out uint dpiX, out uint dpiY);
                            AppendLog($"  MonitorHandle=0x{hMon.ToInt64():X}, dpiX={dpiX}, dpiY={dpiY}, res={res}");
                        }
                        catch (Exception ex)
                        {
                            AppendLog($"  GetDpiForMonitor failed: {ex.Message}");
                        }
                    }
                    else
                    {
                        AppendLog("  MonitorFromPoint returned 0");
                    }
                }
                using (var g = this.CreateGraphics())
                {
                    AppendLog($"Current window Graphics DPI: {g.DpiX}x{g.DpiY}");
                }
                AppendLog("--- Display Info End ---");
            }
            catch (Exception ex)
            {
                AppendLog("LogDisplayInfo error: " + ex.Message);
            }
        }

        private string GetWindowInfo(IntPtr hWnd)
        {
            try
            {
                if (hWnd == IntPtr.Zero) return "hWnd=0";
                int len = GetWindowTextLength(hWnd);
                var sb = new System.Text.StringBuilder(len + 1);
                GetWindowText(hWnd, sb, sb.Capacity);
                var title = sb.ToString();
                var cls = new System.Text.StringBuilder(256);
                GetClassName(hWnd, cls, cls.Capacity);
                return $"hWnd=0x{hWnd.ToInt64():X}, title='{title}', class='{cls}'";
            }
            catch (Exception ex)
            {
                return "GetWindowInfo error: " + ex.Message;
            }
        }

        private string SafeGetMainModuleFileName(System.Diagnostics.Process p)
        {
            try
            {
                return p.MainModule?.FileName ?? "<no main module>";
            }
            catch
            {
                return "<access denied or unknown>";
            }
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
            var colPress = new DataGridViewTextBoxColumn { Name = "PressDuration", HeaderText = "押下時間(ms)", Width = 100 };
            dgvScenario.Columns.Add(colNo);
            dgvScenario.Columns.Add(colDelay);
            dgvScenario.Columns.Add(colPress);
            for (int i = 1; i <= 10; i++)
            {
                dgvScenario.Columns.Add(new DataGridViewTextBoxColumn { Name = $"Pos{i}", HeaderText = $"座標{i}", Width = 120 });
            }
            // enable row drag/drop
            dgvScenario.AllowDrop = true;
            dgvScenario.MouseDown += DgvScenario_MouseDown;
            dgvScenario.DragOver += DgvScenario_DragOver;
            dgvScenario.DragDrop += DgvScenario_DragDrop;
            dgvScenario.CellMouseDown += DgvScenario_CellMouseDown;
            // context menu
            EnsureRowContextMenu();
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

        // Drag/drop and context menu support
        private ContextMenuStrip? _rowContextMenu;
        private int _dragRowIndex = -1;

        private void EnsureRowContextMenu()
        {
            if (_rowContextMenu != null) return;
            _rowContextMenu = new ContextMenuStrip();
            var insertAbove = new ToolStripMenuItem("行を上に挿入");
            var insertBelow = new ToolStripMenuItem("行を下に挿入");
            var delete = new ToolStripMenuItem("行を削除");
            var moveUp = new ToolStripMenuItem("上へ移動");
            var moveDown = new ToolStripMenuItem("下へ移動");
            insertAbove.Click += (s, e) => { if (dgvScenario.CurrentCell != null) InsertBlankRowAt(dgvScenario.CurrentCell.RowIndex); };
            insertBelow.Click += (s, e) => { if (dgvScenario.CurrentCell != null) InsertBlankRowAt(dgvScenario.CurrentCell.RowIndex + 1); };
            delete.Click += (s, e) => { if (dgvScenario.CurrentCell != null) DeleteRowAt(dgvScenario.CurrentCell.RowIndex); };
            moveUp.Click += (s, e) => { if (dgvScenario.CurrentCell != null) MoveRow(dgvScenario.CurrentCell.RowIndex, Math.Max(0, dgvScenario.CurrentCell.RowIndex - 1)); };
            moveDown.Click += (s, e) => { if (dgvScenario.CurrentCell != null) MoveRow(dgvScenario.CurrentCell.RowIndex, dgvScenario.CurrentCell.RowIndex + 1); };
            _rowContextMenu.Items.AddRange(new ToolStripItem[] { insertAbove, insertBelow, delete, moveUp, moveDown });
        }

        private void DgvScenario_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
            {
                try
                {
                    dgvScenario.CurrentCell = dgvScenario.Rows[e.RowIndex].Cells[e.ColumnIndex >= 0 ? e.ColumnIndex : 0];
                    if (_rowContextMenu != null)
                        _rowContextMenu.Show(Cursor.Position);
                }
                catch { }
            }
        }

        private void DgvScenario_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                var info = dgvScenario.HitTest(e.X, e.Y);
                if (info.RowIndex >= 0)
                {
                    _dragRowIndex = info.RowIndex;
                    dgvScenario.DoDragDrop(_dragRowIndex, DragDropEffects.Move);
                }
            }
        }

        private void DgvScenario_DragOver(object? sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void DgvScenario_DragDrop(object? sender, DragEventArgs e)
        {
            var clientPoint = dgvScenario.PointToClient(new System.Drawing.Point(e.X, e.Y));
            var info = dgvScenario.HitTest(clientPoint.X, clientPoint.Y);
            if (info.RowIndex < 0) return;
            if (e.Data != null && e.Data.GetDataPresent(typeof(int)))
            {
                var obj = e.Data.GetData(typeof(int));
                if (obj is int from)
                {
                    int to = info.RowIndex;
                    MoveRow(from, to);
                }
            }
        }

        private void InsertBlankRowAt(int index)
        {
            index = Math.Max(0, Math.Min(index, dgvScenario.Rows.Count));
            var cells = new object[dgvScenario.ColumnCount];
            for (int i = 0; i < cells.Length; i++) cells[i] = string.Empty;
            dgvScenario.Rows.Insert(index, cells);
            RefreshNoColumn();
        }

        private void DeleteRowAt(int index)
        {
            if (index < 0 || index >= dgvScenario.Rows.Count) return;
            if (dgvScenario.Rows[index].IsNewRow) return;
            dgvScenario.Rows.RemoveAt(index);
            RefreshNoColumn();
        }

        private void MoveRow(int fromIndex, int toIndex)
        {
            if (fromIndex < 0 || fromIndex >= dgvScenario.Rows.Count) return;
            if (toIndex < 0) toIndex = 0;
            if (toIndex >= dgvScenario.Rows.Count) toIndex = dgvScenario.Rows.Count - 1;
            if (fromIndex == toIndex) return;
            var row = dgvScenario.Rows[fromIndex];
            var values = new object[dgvScenario.ColumnCount];
            for (int i = 0; i < dgvScenario.ColumnCount; i++)
            {
                var v = row.Cells[i].Value;
                values[i] = v ?? string.Empty;
            }
            // insert copy
            dgvScenario.Rows.Insert(toIndex, values);
            // adjust removal index if inserting before original
            if (fromIndex >= toIndex) fromIndex += 1;
            dgvScenario.Rows.RemoveAt(fromIndex);
            RefreshNoColumn();
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
                var cells = new object[13];
                cells[0] = string.Empty; // NO: will be filled
                cells[1] = s.Delay;
                cells[2] = s.PressDuration;
                for (int i = 0; i < 10; i++)
                {
                    cells[3 + i] = i < s.Positions.Count ? s.Positions[i] : string.Empty;
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
                // read press duration (new column)
                var pdCell = row.Cells[2].Value;
                if (pdCell != null && int.TryParse(pdCell.ToString(), out var pd))
                    step.PressDuration = pd;
                else
                    step.PressDuration = 100;

                for (int i = 0; i < 10; i++)
                {
                    var v = row.Cells[3 + i].Value?.ToString();
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
                _scriptService?.Resume();
                _isPaused = false;
            }
            else
            {
                _script_service_pause();
            }
        }

        private void _script_service_pause()
        {
            _scriptService?.Pause();
            _isPaused = true;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (!(_scriptService?.IsRunning ?? false)) return;
            _scriptService?.Stop();
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
            AppendLog($"OnLoad: mouseHook created, hookId TBD");
            LogDisplayInfo();
        }

        private void BtnToggleCapture_CheckedChanged(object? sender, EventArgs e)
        {
            _captureMode = btnToggleCapture.Checked;
            btnToggleCapture.Text = _captureMode ? "座標抽出 ON" : "座標抽出 OFF";
            try
            {
                if (_captureMode)
                {
                    btnToggleCapture.BackColor = System.Drawing.Color.FromArgb(144, 238, 144); // lightgreen
                    btnToggleCapture.ForeColor = System.Drawing.Color.Black;
                }
                else
                {
                    btnToggleCapture.BackColor = System.Drawing.SystemColors.Control;
                    btnToggleCapture.ForeColor = System.Drawing.Color.Black;
                }
            }
            catch { }
            if (_captureMode)
            {
                _mouseHook?.Start();
                AppendLog("座標抽出モード: ON");
                LogDisplayInfo();
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
            // ログ出力：マウスフックイベント到達確認
            AppendLog($"OnMouseClick raw: {x},{y}");
            // 座標カラムのセルにフォーカスがある場合のみ抽出
            if (dgvScenario.CurrentCell == null || dgvScenario.CurrentCell.ColumnIndex < 3)
            {
                AppendLog("OnMouseClick: CurrentCell 不在または座標カラムではないため無視");
                return;
            }
            // サブ画面／サブウィンドウ対応：アプリの全フォーム領域を取得して判定
            var pt = new System.Drawing.Point(x, y);
            // クリック位置のウィンドウハンドルを取得し、所有プロセスが自プロセスなら無視する
            var hWnd = WindowFromPoint(pt);
            if (hWnd == IntPtr.Zero)
            {
                AppendLog("WindowFromPoint returned 0 (no window)");
            }
            else
            {
                // ルートウィンドウを取得してプロセスIDを比較する（子ウィンドウやオーバーレイ対策）
                var root = GetAncestor(hWnd, GA_ROOT);
                GetWindowThreadProcessId(root, out uint rootPid);
                AppendLog($"WindowFromPoint hWnd=0x{hWnd.ToInt64():X}, root=0x{root.ToInt64():X}, pid={rootPid}, mypid={(uint)System.Diagnostics.Process.GetCurrentProcess().Id}");
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById((int)rootPid);
                    AppendLog($"  Process: Name={p.ProcessName}, Id={p.Id}, Path={SafeGetMainModuleFileName(p)}");
                }
                catch (Exception ex)
                {
                    AppendLog($"  Process info unavailable: {ex.Message}");
                }
                if (rootPid == (uint)System.Diagnostics.Process.GetCurrentProcess().Id)
                {
                    AppendLog("OnMouseClick: クリック先は自プロセスのウィンドウのため無視");
                    return;
                }
            }

            // 追加判定：Application.OpenForms の領域に含まれるか、
            // クリック先ハンドルが各フォームの子ウィンドウかを確認
            AppendLog($"Probe: pt={pt.X},{pt.Y}, hWnd=0x{hWnd.ToInt64():X}");
            foreach (Form f in Application.OpenForms)
            {
                var formScreenPos = f.PointToScreen(System.Drawing.Point.Empty);
                var rect = new System.Drawing.Rectangle(formScreenPos, f.Size);
                AppendLog($"Form '{f.Name}' rect={rect.X},{rect.Y} size={rect.Width}x{rect.Height} handle=0x{f.Handle.ToInt64():X}");

                // obtain monitor DPI for the monitor where the form is located
                var monForForm = MonitorFromPoint(new System.Drawing.Point(formScreenPos.X + rect.Width / 2, formScreenPos.Y + rect.Height / 2), MONITOR_DEFAULTTONEAREST);
                double formScale = 1.0;
                if (monForForm != IntPtr.Zero)
                {
                    try
                    {
                        var rr = GetDpiForMonitor(monForForm, MDT_EFFECTIVE_DPI, out uint fdpiX, out uint fdpiY);
                        if (rr == 0 && fdpiX > 0) formScale = fdpiX / 96.0;
                        AppendLog($"  form monitor hMon=0x{monForForm.ToInt64():X}, dpiX={fdpiX}, scale={formScale}");
                    }
                    catch (Exception ex)
                    {
                        AppendLog("  GetDpiForMonitor(form) error: " + ex.Message);
                    }
                }

                // compute physical bounds of the form (approx)
                var physRect = new System.Drawing.Rectangle((int)(formScreenPos.X * formScale), (int)(formScreenPos.Y * formScale), (int)(rect.Width * formScale), (int)(rect.Height * formScale));
                var pt_physical = new System.Drawing.Point(x, y);
                var pt_logical_byForm = new System.Drawing.Point((int)(x / formScale), (int)(y / formScale));
                var pt_scaled_byForm = new System.Drawing.Point((int)(x * formScale), (int)(y * formScale));

                // also probe the monitor that contains the click point and try its DPI/scale
                var monForPoint = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
                double ptScale = 1.0;
                if (monForPoint != IntPtr.Zero)
                {
                    try
                    {
                        var r = GetDpiForMonitor(monForPoint, MDT_EFFECTIVE_DPI, out uint pdpiX, out uint pdpiY);
                        if (r == 0 && pdpiX > 0) ptScale = pdpiX / 96.0;
                        AppendLog($"  point monitor hMon=0x{monForPoint.ToInt64():X}, dpiX={pdpiX}, scale={ptScale}");
                    }
                    catch (Exception ex)
                    {
                        AppendLog("  GetDpiForMonitor(point) error: " + ex.Message);
                    }
                }
                var pt_logical_byPoint = new System.Drawing.Point((int)(x / ptScale), (int)(y / ptScale));
                var pt_scaled_byPoint = new System.Drawing.Point((int)(x * ptScale), (int)(y * ptScale));

                bool match_physical = physRect.Contains(pt_physical);
                bool match_logical_byForm = rect.Contains(pt_logical_byForm);
                bool match_scaled_byForm = rect.Contains(pt_scaled_byForm);
                bool match_logical_byPoint = rect.Contains(pt_logical_byPoint);
                bool match_scaled_byPoint = rect.Contains(pt_scaled_byPoint);

                AppendLog($"  matches: physical={match_physical}, logical_byForm={match_logical_byForm}, scaled_byForm={match_scaled_byForm}, logical_byPoint={match_logical_byPoint}, scaled_byPoint={match_scaled_byPoint}, physRect={physRect}");

                if (match_physical || match_logical_byForm || match_scaled_byForm || match_logical_byPoint || match_scaled_byPoint)
                {
                    AppendLog($"Point considered inside form '{f.Name}' by one of transforms -> ignore");
                    return;
                }

                // 親チェーンを辿ってクリック先ハンドルがこのフォームの子ウィンドウか確認
                var cur = hWnd;
                while (cur != IntPtr.Zero)
                {
                    AppendLog($" ancestor=0x{cur.ToInt64():X}, info={GetWindowInfo(cur)}");
                    if (cur == f.Handle)
                    {
                        AppendLog($"hWnd is child of form '{f.Name}' -> ignore");
                        return;
                    }
                    var parent = GetParent(cur);
                    if (parent == IntPtr.Zero)
                    {
                        var root = GetAncestor(cur, GA_ROOT);
                        if (root != IntPtr.Zero && root == f.Handle)
                        {
                            AppendLog($"ancestor root==form '{f.Name}' -> ignore");
                            return;
                        }
                        break;
                    }
                    cur = parent;
                }
            }
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
            // 座標抽出後にアプリを前面に戻す
            // 最前面化を安定させるため、詳細ログとリトライを行う
            try
            {
                const uint ASFW_ANY = 0xFFFFFFFF;
                var fgWnd = GetForegroundWindow();
                uint fgThread = GetWindowThreadProcessId(fgWnd, out uint fgPid);
                uint currentThread = GetCurrentThreadId();
                AppendLog($"Foreground hWnd=0x{fgWnd.ToInt64():X}, fgPid={fgPid}, fgThread={fgThread}, currentThread={currentThread}");

                // AllowSetForegroundWindow を試す（診断目的）
                try
                {
                    var allowed = AllowSetForegroundWindow(ASFW_ANY);
                    AppendLog($"AllowSetForegroundWindow(ASFW_ANY)={allowed}, err={Marshal.GetLastWin32Error()}");
                }
                catch (Exception ex)
                {
                    AppendLog("AllowSetForegroundWindow failed: " + ex.Message);
                }

                bool success = false;
                for (int attempt = 0; attempt < 5 && !success; attempt++)
                {
                    AppendLog($"Foreground attempt {attempt + 1}");
                    // Attach current thread to the foreground window thread, so SetForegroundWindow has a better chance
                    bool attached = false;
                    try
                    {
                        attached = AttachThreadInput(currentThread, fgThread, true);
                        AppendLog($"AttachThreadInput current->{fgThread} result={attached}, err={Marshal.GetLastWin32Error()}");
                    }
                    catch (Exception ex)
                    {
                        AppendLog("AttachThreadInput exception: " + ex.Message);
                    }

                    // Try multiple ways to bring window up
                    try { ShowWindowAsync(this.Handle, SW_SHOWNORMAL); } catch { }
                    try { BringWindowToTop(this.Handle); } catch { }
                    try { SetWindowPos(this.Handle, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE); } catch { }
                    try { SetForegroundWindow(this.Handle); } catch { }
                    try { SetActiveWindow(this.Handle); } catch { }

                    // simulate ALT to help focus rules
                    try
                    {
                        keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
                        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                    }
                    catch { }

                    // detach if we attached
                    if (attached)
                    {
                        try { AttachThreadInput(currentThread, fgThread, false); } catch { }
                    }

                    // check if we're now foreground
                    var nowFg = GetForegroundWindow();
                    if (nowFg == this.Handle)
                    {
                        AppendLog("Became foreground window");
                        success = true;
                        break;
                    }

                    AppendLog($"Not foreground yet (nowFg=0x{nowFg.ToInt64():X}), sleeping before retry");
                    Thread.Sleep(60);
                }
                AppendLog($"Foreground attempts finished, success={success}");
            }
            catch (Exception ex)
            {
                AppendLog("Foreground helper failed: " + ex.Message);
            }

            // If previous attempts failed to make us foreground, try a temporary focus-grabber window trick
            try
            {
                // if still not foreground, create a tiny topmost window, activate it, then close it
                var nowFg2 = GetForegroundWindow();
                if (nowFg2 != this.Handle)
                {
                    AppendLog("Attempting temporary focus grabber trick");
                    TemporaryFocusGrabber();
                    var nowFg3 = GetForegroundWindow();
                    AppendLog($"After grabber, foreground hWnd=0x{nowFg3.ToInt64():X}");
                    // final attempt
                    try { SetForegroundWindow(this.Handle); } catch { }
                }
            }
            catch (Exception ex)
            {
                AppendLog("Focus grabber failed: " + ex.Message);
            }
            // make topmost briefly and keep it for a short duration to ensure visibility
            var ok1 = SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            var err1 = Marshal.GetLastWin32Error();
            AppendLog($"SetWindowPos TOPMOST result={ok1}, err={err1}");
            try { Thread.Sleep(220); } catch { }
            var ok2 = SetWindowPos(this.Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
            var err2 = Marshal.GetLastWin32Error();
            AppendLog($"SetWindowPos NOTOPMOST result={ok2}, err={err2}");
            // 追加フォールバック：ShowWindow + SetWindowPos(HWND_TOP) + Altキー送信
            try
            {
                ShowWindow(this.Handle, SW_SHOWNORMAL);
                SetWindowPos(this.Handle, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                // simulate ALT press/release to help SetForegroundWindow permissions
                keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                SetForegroundWindow(this.Handle);
                SetActiveWindow(this.Handle);
                AppendLog("Fallback foreground attempts executed");
            }
            catch (Exception ex)
            {
                AppendLog("Fallback foreground failed: " + ex.Message);
            }
            this.BringToFront();
            this.Activate();
            // DataGridViewの選択セルにフォーカス・編集状態・選択状態を明示的に戻す
            BeginInvoke(new Action(() => {
                dgvScenario.CurrentCell = dgvScenario.Rows[rowIndex].Cells[targetCol];
                dgvScenario.Select();
                dgvScenario.Focus();
                if (dgvScenario.CurrentCell != null)
                {
                    dgvScenario.BeginEdit(false);
                    dgvScenario.CurrentCell.Selected = true;
                    AppendLog($"選択セルにフォーカス復帰: row={rowIndex}, col={targetCol}");
                }
            }));
        }

        // Temporary tiny topmost window trick to nudge the OS focus rules
        private void TemporaryFocusGrabber()
        {
            Form grab = new Form();
            try
            {
                grab.FormBorderStyle = FormBorderStyle.None;
                grab.ShowInTaskbar = false;
                grab.StartPosition = FormStartPosition.Manual;
                // place it near the top-left of main window
                var loc = this.PointToScreen(System.Drawing.Point.Empty);
                grab.Location = new System.Drawing.Point(loc.X + 8, loc.Y + 8);
                grab.Size = new System.Drawing.Size(2, 2);
                grab.TopMost = true;
                grab.Opacity = 0; // invisible
                // show and activate briefly
                grab.Show();
                grab.BringToFront();
                grab.Activate();
                Application.DoEvents();
                Thread.Sleep(60);
            }
            catch (Exception ex)
            {
                AppendLog("TemporaryFocusGrabber exception: " + ex.Message);
            }
            finally
            {
                try { grab.Close(); grab.Dispose(); } catch { }
            }
        }
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

}
