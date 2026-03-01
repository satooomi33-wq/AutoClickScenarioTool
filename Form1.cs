using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using AutoClickScenarioTool.Models;
using System.Text.Json;
using AutoClickScenarioTool.Services;
using System.IO.Ports;
namespace AutoClickScenarioTool
{
    public partial class Form1 : Form
    {
        // Serial/Teensy support
        private SerialPort? _teensyPort;
        private CancellationTokenSource? _serialScanCts;
        private readonly Dictionary<string, string> _detectedTeensies = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private string _selectedOutputTarget = "PC"; // "PC" or COM port name
        // ComboBox for selecting focus target app before playback (declared in Designer)
        private const string NoFocusLabel = "アプリフォーカス無";
        private enum CaptureModeState { Disabled, Mouse, Key }
        private CaptureModeState _captureModeState = CaptureModeState.Disabled;
        // Designer-managed ToolStrip controls are used instead of runtime-created ones.
        private GlobalKeyboardHook? _keyboardHook;
        private DateTime _lastKeyTime = DateTime.MinValue;
        private string _lastKeyCaptured = string.Empty;
        private readonly DataService _dataService = new DataService();
        private readonly InputService _inputService = new InputService();
        private Models.DefaultSettings _defaultSettings = new Models.DefaultSettings();
        private ScriptService? _scriptService;

        private bool _isPaused = false;
        private DataGridViewEditMode _savedEditMode = DataGridViewEditMode.EditOnKeystroke;
        // undo/redo stacks store serialized grid state (JSON of ScenarioStep list)
        private readonly List<string> _undoStack = new List<string>();
        private readonly List<string> _redoStack = new List<string>();
        private const int MaxHistory = 5;

        // Expose script running state for external helpers (e.g. message filter)
        public bool IsScriptRunning => _scriptService?.IsRunning ?? false;

        

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
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private void DgvScenario_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            try { PushHistory(); UpdateUndoRedoButtons(); } catch { }
        }

        // History management for undo/redo
        private bool _suppressHistory = false;
        private void ResetHistory()
        {
            _undoStack.Clear();
            _redoStack.Clear();
            // push current state as baseline
            try { var s = JsonSerializer.Serialize(ReadStepsFromGrid()); _undoStack.Add(s); } catch { }
            UpdateUndoRedoButtons();
        }

        private void PushHistory()
        {
            if (_suppressHistory) return;
            try
            {
                var s = JsonSerializer.Serialize(ReadStepsFromGrid());
                if (_undoStack.Count > 0 && _undoStack.Last() == s) return;
                _undoStack.Add(s);
                if (_undoStack.Count > MaxHistory) _undoStack.RemoveAt(0);
                _redoStack.Clear();
                UpdateUndoRedoButtons();
            }
            catch { }
        }

        private void UpdateUndoRedoButtons()
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(UpdateUndoRedoButtons));
                    return;
                }
                try { btnUndo.Enabled = _undoStack.Count > 1; } catch { }
                try { btnRedo.Enabled = _redoStack.Count > 0; } catch { }
            }
            catch { }
        }

        private void ApplyState(string state)
        {
            try
            {
                var list = JsonSerializer.Deserialize<List<ScenarioStep>>(state) ?? new List<ScenarioStep>();
                _suppressHistory = true;
                try { FillGridWithSteps(list); } catch { }
                _suppressHistory = false;
            }
            catch { }
        }

        private void btnUndo_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_undoStack.Count <= 1) return; // nothing to undo
                var current = _undoStack.Last();
                _undoStack.RemoveAt(_undoStack.Count - 1);
                _redoStack.Add(current);
                var prev = _undoStack.LastOrDefault();
                if (prev != null) ApplyState(prev);
                UpdateUndoRedoButtons();
            }
            catch { }
        }

        private void btnRedo_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_redoStack.Count == 0) return;
                var s = _redoStack.Last();
                _redoStack.RemoveAt(_redoStack.Count - 1);
                _undoStack.Add(s);
                ApplyState(s);
                UpdateUndoRedoButtons();
            }
            catch { }
        }

        // Auto-adjust grid columns to fit contents
        private void AdjustGridColumnWidths()
        {
            try
            {
                if (dgvScenario == null) return;
                dgvScenario.SuspendLayout();
                foreach (DataGridViewColumn col in dgvScenario.Columns)
                {
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }
                dgvScenario.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgvScenario.ResumeLayout();
            }
            catch { }
        }

        // Populate the focus-app combo with currently visible top-level windows
        private void PopulateFocusAppList()
        {
            try
            {
                if (cmbFocusApp == null) return;
                var prev = cmbFocusApp.SelectedItem?.ToString() ?? string.Empty;
                cmbFocusApp.Items.Clear();
                cmbFocusApp.Items.Add(NoFocusLabel);
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in Process.GetProcesses())
                {
                    try
                    {
                        if (p.Id == Process.GetCurrentProcess().Id) continue;
                        var h = p.MainWindowHandle;
                        var title = p.MainWindowTitle ?? string.Empty;
                        if (h == IntPtr.Zero) continue;
                        if (string.IsNullOrWhiteSpace(title)) continue;
                        var display = $"{p.ProcessName} — {title}";
                        if (seen.Add(display)) cmbFocusApp.Items.Add(display);
                    }
                    catch { }
            try { AdjustGridColumnWidths(); } catch { }
                }
                // restore previous selection if still present
                if (!string.IsNullOrWhiteSpace(prev))
                {
                    for (int i = 0; i < cmbFocusApp.Items.Count; i++)
                    {
                        if (string.Equals(cmbFocusApp.Items[i]?.ToString(), prev, StringComparison.OrdinalIgnoreCase))
                        {
                            cmbFocusApp.SelectedIndex = i;
                            return;
                        }
                    }
                }
                cmbFocusApp.SelectedIndex = 0;
            }
            catch { }
        }

        // Try to focus the application represented by the selected display string
        private bool TryFocusSelectedApp(string display)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(display) || display == NoFocusLabel) return true;
                foreach (var p in Process.GetProcesses())
                {
                    try
                    {
                        var title = p.MainWindowTitle ?? string.Empty;
                        if (p.Id == Process.GetCurrentProcess().Id) continue;
                        if (p.MainWindowHandle == IntPtr.Zero) continue;
                        var d = $"{p.ProcessName} — {title}";
                        if (string.Equals(d, display, StringComparison.OrdinalIgnoreCase))
                        {
                            var hwnd = p.MainWindowHandle;
                            if (hwnd == IntPtr.Zero) return false;
                            return TryBringWindowToFront(hwnd);
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return false;
        }

        // ---- Serial monitor helpers ----
        private CancellationTokenSource? _serialMonitorCts2;

        private void RestartSerialMonitorIfNeeded()
        {
            try
            {
                StopSerialMonitor();
                if (!string.Equals(_selectedOutputTarget, "PC", StringComparison.OrdinalIgnoreCase))
                {
                    StartSerialMonitor();
                }
            }
            catch { }
        }

        private void StartSerialMonitor()
        {
            try
            {
                StopSerialMonitor();
                _serialMonitorCts2 = new CancellationTokenSource();
                var token = _serialMonitorCts2.Token;
                Task.Run(async () =>
                {
                    while (!token.IsCancellationRequested)
                    {
                        try
                        {
                            if (!string.Equals(_selectedOutputTarget, "PC", StringComparison.OrdinalIgnoreCase))
                            {
                                // ensure port is open
                                if (_teensyPort == null || !_teensyPort.IsOpen)
                                {
                                    try
                                    {
                                        if (_teensyPort != null) { _teensyPort.Dispose(); _teensyPort = null; }
                                        _teensyPort = new SerialPort(_selectedOutputTarget, 115200) { NewLine = "\n", ReadTimeout = 500, WriteTimeout = 500 };
                                        _teensyPort.DataReceived += TeensyPort_DataReceived;
                                        _teensyPort.Open();
                                    }
                                    catch { try { _teensyPort = null; } catch { } }
                                }

                                // update UI status
                                if (_teensyPort != null && _teensyPort.IsOpen)
                                {
                                    SetSerialStatus($"接続: {_teensyPort.PortName}");
                                }
                                else
                                {
                                    SetSerialStatus("未接続");
                                    // try reconnect if enabled
                                    if (!chkAutoReconnect.Checked)
                                        break;
                                }
                            }
                            else
                            {
                                SetSerialStatus("未使用(PC)");
                                break;
                            }
                        }
                        catch { }
                        await Task.Delay(1500, token).ConfigureAwait(false);
                    }
                }, token);
            }
            catch { }
        }

        private void StopSerialMonitor()
        {
            try
            {
                try { _serialMonitorCts2?.Cancel(); } catch { }
                try { _serialMonitorCts2?.Dispose(); } catch { }
                _serialMonitorCts2 = null;
            }
            catch { }
        }

        private void CloseTeensyPort()
        {
            try
            {
                if (_teensyPort != null)
                {
                    try { _teensyPort.DataReceived -= TeensyPort_DataReceived; } catch { }
                    try { if (_teensyPort.IsOpen) _teensyPort.Close(); } catch { }
                    try { _teensyPort.Dispose(); } catch { }
                    _teensyPort = null;
                }
            }
            catch { }
        }

        private void TeensyPort_DataReceived(object? sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                var sp = sender as SerialPort;
                if (sp == null) return;
                string line = string.Empty;
                try { line = sp.ReadLine(); } catch { }
                if (string.IsNullOrWhiteSpace(line)) return;
                // append to UI log
                AppendLog($"[Teensy] {line.Trim()}");
            }
            catch { }
        }

        private void SetSerialStatus(string s)
        {
            try
            {
                if (lblSerialStatus == null) return;
                if (lblSerialStatus.InvokeRequired)
                {
                    lblSerialStatus.Invoke(new Action(() => lblSerialStatus.Text = s));
                }
                else lblSerialStatus.Text = s;
            }
            catch { }
        }

        // Designer click handler for refresh serial button
        private void btnRefreshSerial_Click(object sender, EventArgs e)
        {
            try { PopulateSerialPorts(); } catch { }
        }

        // Designer DropDown event wrapper
        private void cmbFocusApp_DropDown(object? sender, EventArgs e)
        {
            try { PopulateFocusAppList(); } catch { }
        }

        // Refresh the serial ports list into UI
        private void PopulateSerialPorts()
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(PopulateSerialPorts));
                    return;
                }
                cmbSerialPorts.Items.Clear();
                var ports = SerialPort.GetPortNames().OrderBy(p => p).ToArray();
                foreach (var p in ports) cmbSerialPorts.Items.Add(p);
                if (cmbSerialPorts.Items.Count > 0)
                {
                    try { cmbSerialPorts.SelectedIndex = 0; } catch { }
                }
                // update output target combo: keep PC option, then append Teensy entries if detected
                cmbOutputTarget.Items.Clear();
                cmbOutputTarget.Items.Add("PC (Local)");
                foreach (var p in ports)
                {
                    cmbOutputTarget.Items.Add($"Teensy: {p}");
                }
                if (cmbOutputTarget.Items.Count > 0) cmbOutputTarget.SelectedIndex = 0;
            }
            catch { }
        }

        

        private void cmbOutputTarget_SelectedIndexChanged(object? sender, EventArgs e)
        {
            try
            {
                var sel = cmbOutputTarget.SelectedItem?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(sel)) return;
                if (sel.StartsWith("Teensy:", StringComparison.OrdinalIgnoreCase))
                {
                    var port = sel.Substring(sel.IndexOf(':') + 1).Trim();
                    // close existing if different
                    if (_teensyPort != null && !_teensyPort.PortName.Equals(port, StringComparison.OrdinalIgnoreCase))
                    {
                        try { _teensyPort.Close(); } catch { }
                        _teensyPort = null;
                    }
                    // open port
                    try
                    {
                        if (_teensyPort == null)
                        {
                            _teensyPort = new SerialPort(port, 115200) { NewLine = "\n", ReadTimeout = 300, WriteTimeout = 300 };
                            _teensyPort.Open();
                        }
                        _selectedOutputTarget = port;
                        AppendLog($"出力先: Teensy ({port}) に設定");
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"Teensy 接続失敗: {ex.Message}");
                        _selectedOutputTarget = "PC";
                        cmbOutputTarget.SelectedIndex = 0;
                    }
                }
                else
                {
                    _selectedOutputTarget = "PC";
                    // close serial if open
                    try { CloseTeensyPort(); } catch { }
                    AppendLog("出力先: PC (Local) に設定");
                }
            }
            catch { }
            // start/stop monitor depending on selection
            try { RestartSerialMonitorIfNeeded(); } catch { }
        }

        // Attempt to bring a window to the foreground. Returns true on success.
        private bool TryBringWindowToFront(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return false;
            try
            {
                const uint ASFW_ANY = 0xFFFFFFFF;
                var fg = GetForegroundWindow();
                uint fgThread = GetWindowThreadProcessId(fg, out uint fgPid);
                uint targetThread = GetWindowThreadProcessId(hWnd, out uint targetPid);

                for (int attempt = 0; attempt < 5; attempt++)
                {
                    bool attached = false;
                    try
                    {
                        // Attach the foreground thread and target thread so SetForegroundWindow has permission
                        attached = AttachThreadInput(fgThread, targetThread, true);
                    }
                    catch { attached = false; }

                    try { ShowWindowAsync(hWnd, SW_SHOWNORMAL); } catch { }
                    try { SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE); } catch { }
                    try { SetForegroundWindow(hWnd); } catch { }
                    try { SetActiveWindow(hWnd); } catch { }

                    try
                    {
                        // simulate ALT press/release to help focus rules
                        keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
                        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                    }
                    catch { }

                    if (attached)
                    {
                        try { AttachThreadInput(fgThread, targetThread, false); } catch { }
                    }

                    var nowFg = GetForegroundWindow();
                    if (nowFg == hWnd) return true;
                    try { Thread.Sleep(80); } catch { }
                }
            }
            catch { }
            return false;
        }

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
            // initialize focus-app list from Designer control
            try { PopulateFocusAppList(); } catch { }

            _script_service_init();

            CreateGridColumns();

            // ScanCode トグルの初期ハンドラ（Designer 上に tsbScanCode を配置してください）
            try
            {
                // Prefer designer field, but also attach to ToolStrip items if present
                if (tsbScanCode != null)
                {
                    tsbScanCode.Click -= TsbScanCode_ClickInternal;
                    tsbScanCode.Click += TsbScanCode_ClickInternal;
                }
                else
                {
                    // fallback: find by name in captureToolStrip items
                    var tsb = captureToolStrip?.Items.OfType<ToolStripButton>().FirstOrDefault(x => string.Equals(x.Name, "tsbScanCode", StringComparison.OrdinalIgnoreCase));
                    if (tsb != null)
                    {
                        tsb.Click -= TsbScanCode_ClickInternal;
                        tsb.Click += TsbScanCode_ClickInternal;
                        // keep designer field in sync if available later
                        try { tsbScanCode = tsb; } catch { }
                    }
                }
            }
            catch { }

            // local handler method to keep code compact
            void TsbScanCode_ClickInternal(object? s, EventArgs e)
            {
                try
                {
                    if (tsbScanCode == null) return;
                    var isChecked = tsbScanCode.Checked; // CheckOnClick toggles Checked before Click
                    if (_scriptService != null)
                        _scriptService.UseScanCode = isChecked;
                    if (_defaultSettings != null)
                        _defaultSettings.UseScanCode = isChecked;
                    AppendLog($"SC モード: {(isChecked ? "ON" : "OFF")}");
                }
                catch { }
            }

            // 保存：現在の EditMode を記憶（復元用）
            try { _savedEditMode = dgvScenario.EditMode; } catch { }
            // 選択（青）状態でも Back/Delete でクリア、ダブルクリックで編集開始できるようにする
            try
            {
                dgvScenario.KeyDown += DgvScenario_KeyDown;
                dgvScenario.CellDoubleClick += DgvScenario_CellDoubleClick;
                dgvScenario.CellEndEdit += DgvScenario_CellEndEdit;
            }
            catch { }

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
            _ = LoadAndApplyDefaultsAsync();
            try { PopulateSerialPorts(); } catch { }
            // ensure there's at least one editable row on startup
            try { EnsureGridHasRow(); } catch { }

            // defer positioning until the form is shown so layout/anchors are applied
            this.Shown += Form1_Shown;
        }

        private void Form1_Shown(object? sender, EventArgs e)
        {
            // Designer layout should be used as-is; skip runtime repositioning to
            // ensure runtime matches the VS Designer view.
            // (Previously this called PositionControlsRelativeToFileSelector())

            // Ensure form is visible on primary screen. If it's mostly off-screen,
            // move it to the primary screen center (helps when user closed on other monitor).
            try
            {
                var primary = Screen.PrimaryScreen?.WorkingArea;
                if (primary != null)
                {
                    var inter = Rectangle.Intersect(this.Bounds, primary.Value);
                    if (inter.Width < 50 || inter.Height < 50)
                    {
                        int x = primary.Value.Left + Math.Max(0, (primary.Value.Width - this.Width) / 2);
                        int y = primary.Value.Top + 40;
                        this.StartPosition = FormStartPosition.Manual;
                        this.Location = new System.Drawing.Point(x, y);
                    }
                }
            }
            catch { }
        }

        private void PositionControlsRelativeToFileSelector()
        {
            // Positioning logic disabled to preserve Designer coordinates at runtime.
            // If you need automatic centering, re-enable this method.
            return;
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

        private async void btnSaveDefaults_Click(object? sender, EventArgs e)
        {
            try
            {
                if (!int.TryParse(txtDefaultDelay.Text, out int d) || d < 0)
                {
                    MessageBox.Show("遅延は0以上の整数で指定してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (!int.TryParse(txtDefaultPressDuration.Text, out int p) || p < 0)
                {
                    MessageBox.Show("押下時間は0以上の整数で指定してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // parse humanize bounds from textboxes (if invalid, show error)
                int hLower = _defaultSettings?.HumanizeLower ?? 30;
                int hUpper = _defaultSettings?.HumanizeUpper ?? 100;
                if (!string.IsNullOrWhiteSpace(txtHumanizeLower.Text))
                {
                    if (!int.TryParse(txtHumanizeLower.Text, out hLower) || hLower < 0)
                    {
                        MessageBox.Show("擬人化下限は0以上の整数で指定してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(txtHumanizeUpper.Text))
                {
                    if (!int.TryParse(txtHumanizeUpper.Text, out hUpper) || hUpper < 0)
                    {
                        MessageBox.Show("擬人化上限は0以上の整数で指定してください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                if (hUpper < hLower)
                {
                    MessageBox.Show("擬人化の上限は下限以上にしてください", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var settings = new Models.DefaultSettings
                {
                    Delay = d,
                    PressDuration = p,
                    HumanizeEnabled = _defaultSettings?.HumanizeEnabled ?? false,
                    HumanizeLower = hLower,
                    HumanizeUpper = hUpper,
                    UseScanCode = (tsbScanCode != null) ? tsbScanCode.Checked : (_defaultSettings?.UseScanCode ?? false)
                };

                btnSaveDefaults.Enabled = false;
                await _dataService.SaveDefaultsAsync(settings).ConfigureAwait(false);
                _defaultSettings = settings;
                // apply to runtime script service if available
                try
                {
                    if (_scriptService != null)
                    {
                        _scriptService.HumanizeEnabled = _defaultSettings.HumanizeEnabled;
                        _scriptService.HumanizeLower = _defaultSettings.HumanizeLower;
                        _scriptService.HumanizeUpper = _defaultSettings.HumanizeUpper;
                    }
                }
                catch { }
                try
                {
                    Invoke(new Action(() =>
                    {
                        AppendLog("初期値を保存しました");
                        MessageBox.Show("初期値を保存しました", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
                catch { }
            }
            catch (Exception ex)
            {
                AppendLog($"Save defaults failed: {ex.Message}");
                try { Invoke(new Action(() => MessageBox.Show($"初期値の保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error))); } catch { }
            }
            finally
            {
                try { Invoke(new Action(() => btnSaveDefaults.Enabled = true)); } catch { }
            }
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

        // 現在フォアグラウンドのプロセスが Visual Studio (devenv) かどうかを判定
        private bool IsForegroundProcessVisualStudio()
        {
            try
            {
                var fg = GetForegroundWindow();
                if (fg == IntPtr.Zero) return false;
                GetWindowThreadProcessId(fg, out uint pid);
                try
                {
                    var p = Process.GetProcessById((int)pid);
                    var name = (p.ProcessName ?? string.Empty).ToLowerInvariant();
                    return name.Contains("devenv");
                }
                catch { return false; }
            }
            catch { return false; }
        }

        // 親プロセスを辿って起動元が Visual Studio (devenv) かどうか判定（Toolhelp を使用）
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

        private const uint TH32CS_SNAPPROCESS = 0x00000002;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct PROCESSENTRY32
        {
            public uint dwSize;
            public uint cntUsage;
            public uint th32ProcessID;
            public UIntPtr th32DefaultHeapID;
            public uint th32ModuleID;
            public uint cntThreads;
            public uint th32ParentProcessID;
            public int pcPriClassBase;
            public uint dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szExeFile;
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool Process32First(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool Process32Next(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        private bool IsLaunchedFromVisualStudio()
        {
            try
            {
                var snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
                if (snapshot == IntPtr.Zero || snapshot == new IntPtr(-1)) return false;
                try
                {
                    var map = new Dictionary<uint, (uint parent, string exe)>();
                    var entry = new PROCESSENTRY32();
                    entry.dwSize = (uint)Marshal.SizeOf<PROCESSENTRY32>();
                    if (Process32First(snapshot, ref entry))
                    {
                        do
                        {
                            map[entry.th32ProcessID] = (entry.th32ParentProcessID, entry.szExeFile ?? string.Empty);
                        } while (Process32Next(snapshot, ref entry));
                    }

                    uint cur = (uint)Process.GetCurrentProcess().Id;
                    var visited = new HashSet<uint>();
                    while (cur != 0 && !visited.Contains(cur))
                    {
                        visited.Add(cur);
                        if (!map.TryGetValue(cur, out var info)) break;
                        var parent = info.parent;
                        if (parent == 0) break;
                        if (map.TryGetValue(parent, out var pinfo))
                        {
                            var name = (pinfo.exe ?? string.Empty).ToLowerInvariant();
                            if (name.Contains("devenv")) return true;
                        }
                        cur = parent;
                    }
                }
                finally
                {
                    try { CloseHandle(snapshot); } catch { }
                }
            }
            catch { }
            return false;
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
            // Ensure Invoke is called with the parameter so delegate parameter counts match
            _scriptService.OnPaused += (idx) => Invoke(new Action<int>(HandlePaused), idx);
            // apply default scan-code setting if present
            try
            {
                if (_defaultSettings != null && _scriptService != null)
                    _scriptService.UseScanCode = _defaultSettings.UseScanCode;
            }
            catch { }
            // wire external send override to route key sends to Teensy when selected
            try
            {
                if (_scriptService != null)
                {
                    _scriptService.ExternalSendOverride = (keySpec, duration, useScan) =>
                    {
                        try
                        {
                            // if user selected PC, do not override; let ScriptService fallback to InputService
                            if (string.Equals(_selectedOutputTarget, "PC", StringComparison.OrdinalIgnoreCase))
                                return false;

                            // otherwise, attempt to send to Teensy serial port
                            if (!string.IsNullOrWhiteSpace(_selectedOutputTarget) && _teensyPort != null && _teensyPort.IsOpen)
                            {
                                try
                                {
                                    var cmd = $"KEY:{keySpec}:{duration}";
                                    _teensyPort.WriteLine(cmd);
                                    return true;
                                }
                                catch { return false; }
                            }
                        }
                        catch { }
                        return false;
                    };
                }
            }
            catch { }
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
                dgvScenario.Columns.Add(new DataGridViewTextBoxColumn { Name = $"Action{i}", HeaderText = $"アクション{i}", Width = 200 });
            }
            // enable row drag/drop
            dgvScenario.AllowDrop = true;
            dgvScenario.MouseDown += DgvScenario_MouseDown;
            dgvScenario.DragOver += DgvScenario_DragOver;
            dgvScenario.DragDrop += DgvScenario_DragDrop;
            dgvScenario.CellMouseDown += DgvScenario_CellMouseDown;
            dgvScenario.CellValidating += DgvScenario_CellValidating;
            dgvScenario.RowsAdded += dgvScenario_RowsChanged;
            // also handle rows-added specifically to populate defaults
            dgvScenario.RowsAdded += DgvScenario_RowsAdded;
            // context menu
            EnsureRowContextMenu();
        }

        private void DgvScenario_RowsAdded(object? sender, DataGridViewRowsAddedEventArgs e)
        {
            // When rows are added (user or programmatically), fill default Delay/PressDuration
            try
            {
                for (int i = e.RowIndex; i < e.RowIndex + e.RowCount && i < dgvScenario.Rows.Count; i++)
                {
                    var row = dgvScenario.Rows[i];
                    if (row.IsNewRow) continue;
                    if (_defaultSettings != null)
                    {
                        // only set if cells empty (use safe null propagation)
                        if (dgvScenario.ColumnCount > 1)
                        {
                            var v1 = row.Cells[1].Value?.ToString();
                            if (string.IsNullOrWhiteSpace(v1))
                                row.Cells[1].Value = _defaultSettings.Delay;
                        }
                        if (dgvScenario.ColumnCount > 2)
                        {
                            var v2 = row.Cells[2].Value?.ToString();
                            if (string.IsNullOrWhiteSpace(v2))
                                row.Cells[2].Value = _defaultSettings.PressDuration;
                        }
                    }
                }
            }
            catch { }
        }

        private void AddBlankRowWithDefaults()
        {
            try
            {
                var cells = new object[dgvScenario.ColumnCount];
                cells[0] = string.Empty;
                cells[1] = _defaultSettings?.Delay ?? 500;
                cells[2] = _defaultSettings?.PressDuration ?? 100;
                for (int i = 3; i < dgvScenario.ColumnCount; i++) cells[i] = string.Empty;
                dgvScenario.Rows.Add(cells);
                RefreshNoColumn();
            try { AdjustGridColumnWidths(); } catch { }
            }
            catch { }
        }

        private void EnsureGridHasRow()
        {
            try
            {
                // If grid has no non-new rows, add one blank row
                bool hasNonNew = false;
                foreach (DataGridViewRow r in dgvScenario.Rows)
                {
                    if (!r.IsNewRow) { hasNonNew = true; break; }
                }
                if (!hasNonNew)
                {
                    AddBlankRowWithDefaults();
                }
            }
            catch { }
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

        // CreateCaptureToolbar removed: using designer-added ToolStrip and buttons.

        private void UpdateToolbarButtons()
        {
            try
            {
                if (tsbDisable == null || tsbMouse == null || tsbKey == null || tslCaptureStatus == null) return;
                tsbDisable.Checked = _captureModeState == CaptureModeState.Disabled;
                tsbMouse.Checked = _captureModeState == CaptureModeState.Mouse;
                tsbKey.Checked = _captureModeState == CaptureModeState.Key;
                tslCaptureStatus.Text = $"キャプチャ: {_captureModeState}";
            }
            catch { }
        }

        private void tsbDisable_Click(object sender, EventArgs e)
        {
            SetCaptureMode(CaptureModeState.Disabled);
        }

        private void tsbMouse_Click(object sender, EventArgs e)
        {
            SetCaptureMode(CaptureModeState.Mouse);
        }

        private void tsbKey_Click(object sender, EventArgs e)
        {
            SetCaptureMode(CaptureModeState.Key);
        }

        private void SetCaptureMode(CaptureModeState mode)
        {
            // Ensure any in-progress cell edit is ended and move focus away from the grid
            try
            {
                if (dgvScenario != null && dgvScenario.IsCurrentCellInEditMode)
                {
                    // commit/cancel edit so cells are not left editable when switching modes
                    try { dgvScenario.EndEdit(); } catch { }
                }
                // move focus to the toolbar so DataGridView is not focused (prevents accidental editing)
                try { captureToolStrip?.Focus(); } catch { }
            }
            catch { }

            if (_captureModeState == mode) { UpdateToolbarButtons(); return; }
            try
            {
                if (_captureModeState == CaptureModeState.Mouse)
                    _mouseHook?.Stop();
                else if (_captureModeState == CaptureModeState.Key)
                    _keyboardHook?.Stop();
            }
            catch { }

            _captureModeState = mode;

            if (mode == CaptureModeState.Mouse)
            {
                _mouseHook ??= new GlobalMouseHook();
                _mouseHook.OnMouseClick -= HandleGlobalMouseClick;
                _mouseHook.OnMouseClick += HandleGlobalMouseClick;
                _mouseHook.Start();
                AppendLog("Capture mode: Mouse ON");
            }
            else if (mode == CaptureModeState.Key)
            {
                _keyboardHook ??= new GlobalKeyboardHook();
                _keyboardHook.OnKeyPressed -= HandleGlobalKeyPress;
                _keyboardHook.OnKeyPressed += HandleGlobalKeyPress;
                _keyboardHook.SuppressKeys = true;
                // Only suppress when our form has focus and the selected cell is an Action column
                _keyboardHook.ShouldSuppress = () =>
                {
                    try
                    {
                        if (!this.Focused) return false;
                        if (dgvScenario == null) return false;
                        if (!dgvScenario.Focused) return false;
                        var cell = dgvScenario.CurrentCell;
                        if (cell == null) return false;
                        var col = dgvScenario.Columns[cell.ColumnIndex];
                        if (!col.Name.StartsWith("Action", StringComparison.OrdinalIgnoreCase)) return false;

                        // Allow Backspace/Delete so user can clear cell contents while in capture mode
                        if ((GetAsyncKeyState((int)System.Windows.Forms.Keys.Back) & 0x8000) != 0) return false;
                        if ((GetAsyncKeyState((int)System.Windows.Forms.Keys.Delete) & 0x8000) != 0) return false;

                        return true; // suppress other keys while capturing
                    }
                    catch { return false; }
                };
                // prevent DataGridView from entering edit mode on key press
                try { if (dgvScenario != null) { _savedEditMode = dgvScenario.EditMode; dgvScenario.EditMode = DataGridViewEditMode.EditProgrammatically; } } catch { }
                _keyboardHook.Start();
                AppendLog("Capture mode: Key ON");
            }
            else
            {
                AppendLog("Capture mode: Disabled");
            }

            // restore edit mode when leaving key capture (ensure UI-thread)
            if (mode != CaptureModeState.Key)
            {
                try
                {
                    if (InvokeRequired)
                        Invoke(new Action(() => { try { if (dgvScenario != null) dgvScenario.EditMode = _savedEditMode; } catch { } }));
                    else
                        if (dgvScenario != null) dgvScenario.EditMode = _savedEditMode;
                }
                catch { }
            }

            UpdateToolbarButtons();
        }

        private void HandleGlobalKeyPress(string keyName)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => HandleGlobalKeyPress(keyName)));
                return;
            }
            var now = DateTime.UtcNow;
            if (_lastKeyCaptured == keyName && (now - _lastKeyTime).TotalMilliseconds < 300) return;
            _lastKeyCaptured = keyName; _lastKeyTime = now;

            AppendLog($"OnKeyPress raw: {keyName}");
            if (dgvScenario.CurrentCell == null) { AppendLog("No current cell, ignoring key capture"); return; }
            var col = dgvScenario.Columns[dgvScenario.CurrentCell.ColumnIndex];
            if (!col.Name.StartsWith("Action", StringComparison.OrdinalIgnoreCase)) { AppendLog("Current col not Action*, ignoring"); return; }

            var spec = keyName;
            if ((GetAsyncKeyState((int)System.Windows.Forms.Keys.ControlKey) & 0x8000) != 0) spec = "Ctrl+" + spec;
            if ((GetAsyncKeyState((int)System.Windows.Forms.Keys.ShiftKey) & 0x8000) != 0) spec = "Shift+" + spec;
            if ((GetAsyncKeyState((int)System.Windows.Forms.Keys.Menu) & 0x8000) != 0) spec = "Alt+" + spec;

            try
            {
                // Normalize single-letter mains to uppercase for consistent playback
                var finalSpec = spec;
                var parts = finalSpec.Split('+');
                var main = parts.Last();
                if (main.Length == 1 && char.IsLetter(main[0]))
                {
                    parts[parts.Length - 1] = main.ToUpperInvariant();
                    finalSpec = string.Join("+", parts);
                }

                dgvScenario.CurrentCell.Value = finalSpec;
                // force UI refresh to ensure displayed value matches logged value
                try
                {
                    dgvScenario.InvalidateCell(dgvScenario.CurrentCell);
                    dgvScenario.Refresh();
                    dgvScenario.Update();
                }
                catch { }
                RefreshNoColumn();
                AppendLog($"Captured key -> {finalSpec}");

                // Re-apply the finalSpec shortly after to override any late-arriving lowercase input
                _ = Task.Run(async () =>
                {
                    await Task.Delay(120).ConfigureAwait(false);
                    try
                    {
                        Invoke(new Action(() =>
                        {
                            try
                            {
                                if (dgvScenario.CurrentCell != null)
                                {
                                    var col = dgvScenario.Columns[dgvScenario.CurrentCell.ColumnIndex];
                                    if (col.Name.StartsWith("Action", StringComparison.OrdinalIgnoreCase))
                                    {
                                        dgvScenario.CurrentCell.Value = finalSpec;
                                        dgvScenario.InvalidateCell(dgvScenario.CurrentCell);
                                        dgvScenario.Refresh();
                                    }
                                }
                            }
                            catch { }
                        }));
                    }
                    catch { }
                });
            }
            catch (Exception ex) { AppendLog("Key capture write failed: " + ex.Message); }
        }

        private void DgvScenario_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

                // 右クリック: コンテキストメニュー表示（既存動作）
                if (e.Button == MouseButtons.Right)
                {
                    dgvScenario.CurrentCell = dgvScenario.Rows[e.RowIndex].Cells[e.ColumnIndex >= 0 ? e.ColumnIndex : 0];
                    if (_rowContextMenu != null)
                        _rowContextMenu.Show(Cursor.Position);
                    return;
                }

                // 左クリック: 選択だけでなく編集開始を試みる
                if (e.Button == MouseButtons.Left)
                {
                    dgvScenario.CurrentCell = dgvScenario.Rows[e.RowIndex].Cells[e.ColumnIndex];

                    // 編集開始は「キャプチャ無（Disabled）」時のみ許可する
                    if (_captureModeState == CaptureModeState.Disabled)
                    {
                        var col = dgvScenario.Columns[e.ColumnIndex];
                        var name = col.Name ?? string.Empty;
                        // Delay / PressDuration / Action* の列はマウスで編集開始を許可
                        if (string.Equals(name, "Delay", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(name, "PressDuration", StringComparison.OrdinalIgnoreCase)
                            || name.StartsWith("Action", StringComparison.OrdinalIgnoreCase))
                        {
                            try { dgvScenario.BeginEdit(true); } catch { }
                        }
                    }
                }
            }
            catch { }
        }

        // キー押下時の補助処理：選択状態(編集外)で Back/Delete を押したらセルを空にする
        private void DgvScenario_KeyDown(object? sender, KeyEventArgs e)
        {
            try
            {
                if (e == null) return;

                // キャプチャの Key モード中はグローバルフックが処理するためここでは無視
                if (_captureModeState == CaptureModeState.Key) return;

                if (dgvScenario.CurrentCell == null) return;
                var col = dgvScenario.Columns[dgvScenario.CurrentCell.ColumnIndex];
                if (!col.Name.StartsWith("Action", StringComparison.OrdinalIgnoreCase)) return;

                // 編集モードでなければ特別扱い
                if (!dgvScenario.IsCurrentCellInEditMode)
                {
                    if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
                    {
                        dgvScenario.CurrentCell.Value = string.Empty;
                        RefreshNoColumn();
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                        return;
                    }

                    // それ以外のキーは編集を開始してから DataGridView に処理させる
                    dgvScenario.BeginEdit(true);
                }
            }
            catch { }
        }

        // ダブルクリックで確実に編集開始する（キャプチャ Key モード中は開始しない）
        private void DgvScenario_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var col = dgvScenario.Columns[e.ColumnIndex];
                if (!col.Name.StartsWith("Action", StringComparison.OrdinalIgnoreCase)) return;

                if (_captureModeState != CaptureModeState.Key)
                {
                    dgvScenario.CurrentCell = dgvScenario.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    dgvScenario.BeginEdit(true);
                }
            }
            catch { }
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

        private void DgvScenario_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                var col = dgvScenario.Columns[e.ColumnIndex];
                var raw = e.FormattedValue?.ToString() ?? string.Empty;
                if (col.Name.StartsWith("Action", StringComparison.OrdinalIgnoreCase))
                {
                    // 空欄は許可（アクションなし）
                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        // 空に正規化しておく
                        dgvScenario.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = string.Empty;
                        return;
                    }

                    var (ok, normalized, error) = AutoClickScenarioTool.Services.KeySpecHelper.ValidateAndNormalize(raw);
                    if (!ok)
                    {
                        // show error, clear the cell and allow edit to finish so user isn't stuck in error state
                        try { MessageBox.Show(error ?? "無効な入力です。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                        try { dgvScenario.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = string.Empty; } catch { }
                        e.Cancel = false;
                        return;
                    }

                    // 正常: 正規化値で上書き（例: 全角→半角→大文字化、修飾子順の標準化）
                    try
                    {
                        dgvScenario.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = normalized;
                    }
                    catch { }
                }
            }
            catch { }
        }

        private bool IsValidKeySpec(string spec)
        {
            if (string.IsNullOrWhiteSpace(spec)) return false;
            var parts = spec.Split('+');
            for (int i = 0; i < parts.Length - 1; i++)
            {
                var m = parts[i].Trim();
                if (!string.Equals(m, "Ctrl", StringComparison.OrdinalIgnoreCase) && !string.Equals(m, "Alt", StringComparison.OrdinalIgnoreCase) && !string.Equals(m, "Shift", StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            var main = parts.Last().Trim();
            if (string.IsNullOrEmpty(main)) return false;
            // Allow single letter/digit or single punctuation/symbol (e.g. ; , + - /)
            if (main.Length == 1 && (char.IsLetterOrDigit(main[0]) || char.IsPunctuation(main[0]) || char.IsSymbol(main[0]))) return true;
            if (int.TryParse(main, out _)) return true;
            var up = main.ToUpperInvariant();
            var allowed = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase)
            {
                "ENTER","TAB","ESC","ESCAPE","BACK","BACKSPACE","SPACE",
                "UP","DOWN","LEFT","RIGHT","HOME","END","INSERT","DELETE","PAGEUP","PAGEDOWN","PGUP","PGDN"
            };
            if (allowed.Contains(up)) return true;
            if (up.StartsWith("F") && int.TryParse(up.Substring(1), out int f) && f >= 1 && f <= 24) return true;
            return false;
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
                var prevRaw = cmbFiles?.Text ?? string.Empty;
                var prev = string.IsNullOrWhiteSpace(prevRaw) ? string.Empty : Path.GetFileNameWithoutExtension(prevRaw);
                cmbFiles.Items.Clear();
                if (Directory.Exists(folder))
                {
                    foreach (var f in _data_service_list_files_safe(folder))
                    {
                        var name = Path.GetFileNameWithoutExtension(f);
                        cmbFiles.Items.Add(name);
                    }
                }
                // restore previous selection if possible
                if (!string.IsNullOrWhiteSpace(prev))
                {
                    for (int i = 0; i < cmbFiles.Items.Count; i++)
                    {
                        if (string.Equals(cmbFiles.Items[i]?.ToString(), prev, StringComparison.OrdinalIgnoreCase))
                        {
                            cmbFiles.SelectedIndex = i;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("Error listing files: " + ex.Message);
            }
        }

        // wrapper to safely enumerate files from DataService (keeps null semantics)
        private IEnumerable<string> _data_service_list_files_safe(string folder)
        {
            try
            {
                return _dataService.ListJsonFiles(folder) ?? Enumerable.Empty<string>();
            }
            catch { return Enumerable.Empty<string>(); }
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
            if (string.IsNullOrWhiteSpace(folder)) return;
            if (!fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                fileName += ".json";
            var path = Path.Combine(folder, fileName);
            if (!File.Exists(path))
            {
                // New filename or no existing file: clear grid and ensure one blank row
                try
                {
                    dgvScenario.Rows.Clear();
                    EnsureGridHasRow();
                }
                catch { }
                return;
            }
            try
            {
                var steps = await _data_service_load(path).ConfigureAwait(false);
                Invoke(new Action(() =>
                {
                    FillGridWithSteps(steps);
                    // reset history after loading a file
                    try { ResetHistory(); } catch { }
                }));
                AppendLog("読み込み完了: " + path);
            }
            catch (Exception ex)
            {
                AppendLog("読み込みエラー: " + ex.Message);
            }
        }

        private Task<List<ScenarioStep>> _data_service_load(string path) => _dataService.LoadAsync(path);

        // Save helper that performs file IO off the UI thread and marshals UI updates back to UI thread
        private async Task _data_service_save_ui(string path, List<ScenarioStep> steps)
        {
            // perform save on background thread
            await Task.Run(async () =>
            {
                await _dataService.SaveAsync(path, steps).ConfigureAwait(false);
            }).ConfigureAwait(false);

            // update UI on UI thread
            try
            {
                // update UI on UI thread and select saved file
                var fileBase = Path.GetFileNameWithoutExtension(path);
                if (InvokeRequired)
                {
                    Invoke(new Action(() =>
                    {
                        AppendLog("保存: " + path);
                        RefreshFileList();
                        try { cmbFiles.Text = fileBase; } catch { }
                        try { ResetHistory(); } catch { }
                    }));
                }
                else
                {
                    AppendLog("保存: " + path);
                    RefreshFileList();
                    try { cmbFiles.Text = fileBase; } catch { }
                    try { ResetHistory(); } catch { }
                }
            }
            catch
            {
                // best-effort: if UI update fails, log to textbox if possible
                try { AppendLog("保存完了(ばんしうひようのかいそくにしっぱい): " + path); } catch { }
            }
        }

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
                await _data_service_save_ui(path, steps).ConfigureAwait(false);
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
                    string outv = string.Empty;
                    if (i < s.Positions.Count)
                    {
                        var raw = s.Positions[i] ?? string.Empty;
                        // If raw is numeric-only with decimal or negative, treat as X and convert to "X,0"
                        var numericOnly = System.Text.RegularExpressions.Regex.Match(raw, "^\\s*-?\\d+(?:\\.\\d+)?\\s*$");
                        if (numericOnly.Success && (raw.Contains('.') || raw.TrimStart().StartsWith("-")))
                        {
                            if (double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dv))
                            {
                                outv = $"{(int)Math.Round(dv)},0";
                            }
                        }
                        else
                        {
                            // If it looks like a coordinate, normalize by rounding components
                            if (raw.Contains(','))
                            {
                                var parts = raw.Split(',');
                                if (parts.Length >= 2
                                    && double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var px)
                                    && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var py))
                                {
                                    outv = $"{(int)Math.Round(px)},{(int)Math.Round(py)}";
                                }
                                else
                                {
                                    outv = raw;
                                }
                            }
                            else
                            {
                                outv = raw;
                            }
                        }
                    }
                    cells[3 + i] = outv;
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
                    var raw = row.Cells[3 + i].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(raw)) continue;
                    // If user entered a single numeric with decimal or negative, treat as X and convert to X,0
                    var numericOnly = System.Text.RegularExpressions.Regex.Match(raw, "^\\s*-?\\d+(?:\\.\\d+)?\\s*$");
                    if (numericOnly.Success && (raw.Contains('.') || raw.TrimStart().StartsWith("-")))
                    {
                        if (double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dv))
                        {
                            step.Positions.Add($"{(int)Math.Round(dv)},0");
                            continue;
                        }
                    }

                    var (ok, normalized, error) = AutoClickScenarioTool.Services.KeySpecHelper.ValidateAndNormalize(raw);
                    if (ok && !string.IsNullOrWhiteSpace(normalized))
                    {
                        step.Positions.Add(normalized);
                    }
                    else
                    {
                        // invalid: skip and log for diagnosis
                        try { AppendLog($"読み飛ばし: 行{row.Index + 1} 列{3 + i} の無効なアクション -> {error ?? "不正な入力"}"); } catch { }
                    }
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
                // Before starting, attempt to focus selected external app if user chose one
                try
                {
                    var sel = cmbFocusApp?.SelectedItem?.ToString() ?? NoFocusLabel;
                    if (!string.IsNullOrWhiteSpace(sel) && sel != NoFocusLabel)
                    {
                        AppendLog($"フォーカス移動先: {sel}");
                        var ok = TryFocusSelectedApp(sel);
                        if (!ok)
                        {
                            AppendLog("選択したアプリへのフォーカスに失敗したため、再生を中止します");
                            try { MessageBox.Show("選択したアプリにフォーカスを当てられません。再生を中止します。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                            return;
                        }
                        AppendLog("フォーカス移動成功");
                        // small pause to allow OS focus changes
                        try { Thread.Sleep(80); } catch { }
                    }
                }
                catch (Exception ex)
                {
                    AppendLog("Focus attempt failed: " + ex.Message);
                    return;
                }

                AppendLog($"実行開始: 行{startIndex + 1} から");
                _isPaused = false;
                btnStart.Enabled = false;
                btnPause.Enabled = true;
                btnStop.Enabled = true;

                // Ensure the service is not left in paused state from previous stop/pause
                try { _scriptService?.Resume(); } catch { }

                // Apply current humanize settings (from defaults/UI) to the script service before starting
                try
                {
                    if (_scriptService != null)
                    {
                        _scriptService.HumanizeEnabled = _defaultSettings?.HumanizeEnabled ?? false;
                        int hl = _defaultSettings?.HumanizeLower ?? 30;
                        int hu = _defaultSettings?.HumanizeUpper ?? 100;
                        if (!string.IsNullOrWhiteSpace(txtHumanizeLower.Text) && int.TryParse(txtHumanizeLower.Text, out var tmpL)) hl = tmpL;
                        if (!string.IsNullOrWhiteSpace(txtHumanizeUpper.Text) && int.TryParse(txtHumanizeUpper.Text, out var tmpU)) hu = tmpU;
                        // normalize
                        _scriptService.HumanizeLower = Math.Max(0, Math.Min(hl, hu));
                        _scriptService.HumanizeUpper = Math.Max(0, Math.Max(hl, hu));
                    }
                }
                catch { }

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
            // Do not call Pause after Stop; Stop() will signal the pause event and cancel the token.
            if (dgvScenario.Rows.Count > 0)
            {
                // bring app front when user requests stop
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

        private void dgvScenario_RowsChanged(object? sender, EventArgs e)
        {
            RefreshNoColumn();
            try { PushHistory(); UpdateUndoRedoButtons(); } catch { }
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
            try
            {
                if (_defaultSettings != null && e.Row != null)
                {
                    // Delay column = index 1, PressDuration = index 2
                    if (dgvScenario.ColumnCount > 1)
                        e.Row.Cells[1].Value = _defaultSettings.Delay;
                    if (dgvScenario.ColumnCount > 2)
                        e.Row.Cells[2].Value = _defaultSettings.PressDuration;
                }
            }
            catch { }
            try { PushHistory(); UpdateUndoRedoButtons(); } catch { }
        }

        private bool _captureMode = false;
        private IntPtr _mainHandle;
        private GlobalMouseHook? _mouseHook;
        private Services.KeyboardMonitor? _kbMonitor;
        // 重複定義削除
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            _mainHandle = this.Handle;
            // _mouseClickFilter の初期化は不要（フィールド未定義のため）
            _mouseHook = new GlobalMouseHook();
            _mouseHook.OnMouseClick += HandleGlobalMouseClick;
            // start low-level keyboard monitor to inspect scan codes for diagnostics
            try
            {
                _kbMonitor = new Services.KeyboardMonitor();
                _kbMonitor.OnKey += (vk, scan, flags) =>
                {
                    try
                    {
                        // Only log when SC toggle is on to reduce noise
                        if (InvokeRequired)
                        {
                            Invoke(new Action(() => { if (tsbScanCode != null && tsbScanCode.Checked) AppendLog($"LLHook vk=0x{vk:X}, scan=0x{scan:X}, flags=0x{flags:X}"); }));
                        }
                        else
                        {
                            if (tsbScanCode != null && tsbScanCode.Checked) AppendLog($"LLHook vk=0x{vk:X}, scan=0x{scan:X}, flags=0x{flags:X}");
                        }
                    }
                    catch { }
                };
                _kbMonitor.Start();
            }
            catch { }
            AppendLog($"OnLoad: mouseHook created, hookId TBD");
            LogDisplayInfo();
        }

        private void Form1_FormClosing(object? sender, FormClosingEventArgs e)
        {
            try { _kbMonitor?.Dispose(); } catch { }
        }

        private async Task LoadAndApplyDefaultsAsync()
        {
            try
            {
                var prevDefaults = _defaultSettings ?? new Models.DefaultSettings();
                var d = await _dataService.LoadDefaultsAsync().ConfigureAwait(false);
                _defaultSettings = d ?? new Models.DefaultSettings();
                try
                {
                    Invoke(new Action(() =>
                    {
                        txtDefaultDelay.Text = _defaultSettings.Delay.ToString();
                        txtDefaultPressDuration.Text = _defaultSettings.PressDuration.ToString();
                        // apply humanize defaults to UI
                        txtHumanizeLower.Text = _defaultSettings.HumanizeLower.ToString();
                        txtHumanizeUpper.Text = _defaultSettings.HumanizeUpper.ToString();
                        // update toggle appearance
                        UpdateHumanizeButtonAppearance();
                        // initialize SC toggle from defaults
                        try
                        {
                            if (tsbScanCode != null)
                                tsbScanCode.Checked = _defaultSettings.UseScanCode;
                            if (_scriptService != null)
                                _scriptService.UseScanCode = _defaultSettings.UseScanCode;
                        }
                        catch { }
                        // apply to script service
                        try
                        {
                            if (_scriptService != null)
                            {
                                _scriptService.HumanizeEnabled = _defaultSettings.HumanizeEnabled;
                                _scriptService.HumanizeLower = _defaultSettings.HumanizeLower;
                                _scriptService.HumanizeUpper = _defaultSettings.HumanizeUpper;
                            }
                        }
                        catch { }
                        // Apply defaults to any existing non-new rows.
                        // Replace if cell is empty OR still has the previous in-memory default value (designer-initialized)
                        try
                        {
                            if (dgvScenario != null)
                            {
                                var prevDelayStr = prevDefaults.Delay.ToString();
                                var prevPressStr = prevDefaults.PressDuration.ToString();
                                foreach (DataGridViewRow r in dgvScenario.Rows)
                                {
                                    if (r.IsNewRow) continue;
                                    try
                                    {
                                        if (r.Cells.Count > 1)
                                        {
                                            var dcell = r.Cells[1].Value?.ToString() ?? string.Empty;
                                            if (string.IsNullOrWhiteSpace(dcell) || dcell == prevDelayStr)
                                                r.Cells[1].Value = _defaultSettings.Delay;
                                        }
                                    }
                                    catch { }
                                    try
                                    {
                                        if (r.Cells.Count > 2)
                                        {
                                            var pcell = r.Cells[2].Value?.ToString() ?? string.Empty;
                                            if (string.IsNullOrWhiteSpace(pcell) || pcell == prevPressStr)
                                                r.Cells[2].Value = _defaultSettings.PressDuration;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { }
                    }));
                }
                catch { }
            }
            catch (Exception ex)
            {
                AppendLog($"Load defaults failed: {ex.Message}");
            }
        }

        private void UpdateHumanizeButtonAppearance()
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(UpdateHumanizeButtonAppearance));
                    return;
                }
                if (_defaultSettings != null && _defaultSettings.HumanizeEnabled)
                {
                    btnToggleHumanize.Text = "擬人化: ON";
                    btnToggleHumanize.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    btnToggleHumanize.Text = "擬人化: OFF";
                    btnToggleHumanize.BackColor = System.Drawing.SystemColors.Control;
                }
            }
            catch { }
        }

        private void btnToggleHumanize_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_defaultSettings == null) _defaultSettings = new Models.DefaultSettings();
                bool currently = _defaultSettings.HumanizeEnabled;
                var msg = currently ? "擬人化を無効にしますか?" : "擬人化を有効にしますか?";
                var dr = MessageBox.Show(msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;

                _defaultSettings.HumanizeEnabled = !currently;
                // reflect immediately in runtime service
                try
                {
                    if (_scriptService != null)
                    {
                        _scriptService.HumanizeEnabled = _defaultSettings.HumanizeEnabled;
                    }
                }
                catch { }
                UpdateHumanizeButtonAppearance();
            }
            catch { }
        }

        

        // グローバルマウスクリック時の座標抽出処理
        public void HandleGlobalMouseClick(int x, int y)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => HandleGlobalMouseClick(x, y)));
                return;
            }
            // Ignore mouse hook events generated by our own playback to avoid
            // overwriting grid cells with synthetic click coordinates.
            try
            {
                if (_scriptService != null && _scriptService.IsRunning)
                {
                    AppendLog("HandleGlobalMouseClick: ignored during script run");
                    return;
                }
            }
            catch { }
            // ログ出力：マウスフックイベント到達確認
            AppendLog($"OnMouseClick raw: {x},{y}");
            // 座標カラムのセルにフォーカスがある場合のみ抽出
            if (dgvScenario.CurrentCell == null || dgvScenario.CurrentCell.ColumnIndex < 3)
            {
                AppendLog("OnMouseClick: CurrentCell 不在または座標カラムではないため無視");
                return;
            }
            // サブ画面／サブウィンドウ対応：アプリの全フォーム領域を取得して判定
            // Convert logical cursor point to physical pixels (per-monitor DPI aware) for comparisons
            var phys = _input_service_real().ConvertToPhysicalPoint(x, y);
            var pt = phys; // use 'pt' where existing code expects a Point in physical pixels
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

            // 追加判定：Application.OpenForms の領域に含まれるか、クリック先ハンドルが各フォームの子ウィンドウかを確認
            AppendLog($"Probe: pt={pt.X},{pt.Y}, hWnd=0x{hWnd.ToInt64():X}");
            foreach (Form f in Application.OpenForms)
            {
                try
                {
                    // Try to obtain the real window rectangle in physical screen coordinates
                    if (f.Handle != IntPtr.Zero && GetWindowRect(f.Handle, out RECT wrec))
                    {
                        var winRect = new System.Drawing.Rectangle(wrec.Left, wrec.Top, wrec.Right - wrec.Left, wrec.Bottom - wrec.Top);
                        AppendLog($"Form '{f.Name}' winRect={winRect} handle=0x{f.Handle.ToInt64():X}");
                        if (winRect.Contains(pt))
                        {
                            AppendLog($"Point considered inside form '{f.Name}' (GetWindowRect) -> ignore");
                            return;
                        }
                    }
                    else
                    {
                        // fallback: use PointToScreen (logical coords) but compare with the raw point conservatively
                        // PointToScreen returns logical coords; convert to physical for conservative comparison
                        var formScreenPos = f.PointToScreen(System.Drawing.Point.Empty);
                        var formPhysTopLeft = _input_service_real().ConvertToPhysicalPoint(formScreenPos.X, formScreenPos.Y);
                        var rect = new System.Drawing.Rectangle(formPhysTopLeft, f.Size);
                        AppendLog($"Form '{f.Name}' rect={rect} handle=0x{f.Handle.ToInt64():X}");
                        if (rect.Contains(pt))
                        {
                            AppendLog($"Point considered inside form '{f.Name}' (fallback) -> ignore");
                            return;
                        }
                    }
                    // also check parent chain: if clicked window belongs to this form, ignore
                    var cur = hWnd;
                    while (cur != IntPtr.Zero)
                    {
                        if (cur == f.Handle)
                        {
                            AppendLog($"hWnd is child of form '{f.Name}' -> ignore");
                            return;
                        }
                        var parent = GetParent(cur);
                        if (parent == IntPtr.Zero) break;
                        cur = parent;
                    }
                }
                catch (Exception ex)
                {
                    AppendLog($"Form probe error: {ex.Message}");
                }
            }
            int rowIndex = dgvScenario.CurrentCell.RowIndex;
            int targetCol = dgvScenario.CurrentCell.ColumnIndex;
            if (dgvScenario.IsCurrentCellInEditMode)
                dgvScenario.EndEdit();
            var raw = $"{x},{y}";
            // Try direct parse first (bypass KeySpecHelper) to avoid single-value normalization issues
            bool parsedDirect = false;
            try
            {
                var parts = raw.Split(',');
                if (parts.Length >= 2
                    && double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dx)
                    && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dy))
                {
                    var xi = (int)Math.Round(dx);
                    var yi = (int)Math.Round(dy);
                    var normalized = $"{xi},{yi}";
                    dgvScenario.Rows[rowIndex].Cells[targetCol].Value = normalized;
                    RefreshNoColumn();
                    AppendLog($"座標取得: {normalized}");
                    parsedDirect = true;
                }
            }
            catch { }

            if (!parsedDirect)
            {
                var (ok, normalized, err) = AutoClickScenarioTool.Services.KeySpecHelper.ValidateAndNormalize(raw);
                if (ok && !string.IsNullOrWhiteSpace(normalized))
                {
                    if (!normalized.Contains(","))
                    {
                        AppendLog($"座標取得失敗(単一値検出): {raw} -> normalized='{normalized}'");
                    }
                    else
                    {
                        dgvScenario.Rows[rowIndex].Cells[targetCol].Value = normalized;
                        RefreshNoColumn();
                        AppendLog($"座標取得: {normalized}");
                    }
                }
                else
                {
                    AppendLog($"座標取得失敗(無効形式): {raw} -> {err}");
                }
            }
            if (rowIndex == dgvScenario.Rows.Count - 2 && dgvScenario.AllowUserToAddRows)
            {
                dgvScenario.Rows.Add();
                RefreshNoColumn();
            }
            // 座標抽出後にアプリを前面に戻す
            if (!System.Diagnostics.Debugger.IsAttached && !IsForegroundProcessVisualStudio())
            {
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

                        // Try multiple ways to bring our main window up
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
                        try { Thread.Sleep(80); } catch { }
                    }
                    AppendLog($"Foreground attempts finished, success={success}");
                }
                catch (Exception ex)
                {
                    AppendLog("Foreground helper failed: " + ex.Message);
                }
            }
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
            // Ignore while script running to avoid capturing our own playback clicks
            try { if (_form.IsScriptRunning) return false; } catch { }
            // 条件を緩和：どこをクリックしても抽出する
            if (_form.dgvScenario.CurrentCell == null) return false;
            int rowIndex = _form.dgvScenario.CurrentCell.RowIndex;
            int targetCol = _form.dgvScenario.CurrentCell.ColumnIndex;
            if (_form.dgvScenario.IsCurrentCellInEditMode)
                _form.dgvScenario.EndEdit();
            var raw = $"{cursorPos.X},{cursorPos.Y}";
            bool parsedDirect = false;
            try
            {
                var parts = raw.Split(',');
                if (parts.Length >= 2
                    && double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dx)
                    && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dy))
                {
                    var xi = (int)Math.Round(dx);
                    var yi = (int)Math.Round(dy);
                    var normalized = $"{xi},{yi}";
                    _form.dgvScenario.Rows[rowIndex].Cells[targetCol].Value = normalized;
                    _form.RefreshNoColumn();
                    _form.AppendLog($"座標取得: {normalized}");
                    parsedDirect = true;
                }
            }
            catch { }

            if (!parsedDirect)
            {
                var (ok, normalized, err) = AutoClickScenarioTool.Services.KeySpecHelper.ValidateAndNormalize(raw);
                if (ok && !string.IsNullOrWhiteSpace(normalized))
                {
                    if (!normalized.Contains(","))
                    {
                        _form.AppendLog($"座標取得失敗(単一値検出): {raw} -> normalized='{normalized}'");
                    }
                    else
                    {
                        _form.dgvScenario.Rows[rowIndex].Cells[targetCol].Value = normalized;
                        _form.RefreshNoColumn();
                        _form.AppendLog($"座標取得: {normalized}");
                    }
                }
                else
                {
                    _form.AppendLog($"座標取得失敗(無効形式): {raw} -> {err}");
                }
            }
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
