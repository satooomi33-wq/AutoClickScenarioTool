using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AutoClickScenarioTool.Models;

namespace AutoClickScenarioTool.Services
{
    public class ScriptService
    {
        // Optional override to send key actions externally (e.g., to Teensy via serial).
        // Signature: (keySpec, durationMs, useScanCode) => returns true if handled and default sending should be skipped.
        public Func<string, int, bool, bool>? ExternalSendOverride { get; set; }

        private readonly InputService _input;
        private CancellationTokenSource? _cts;
        private ManualResetEventSlim _pauseEvent = new ManualResetEventSlim(true);
        public bool IsRunning { get; private set; }
        public int CurrentIndex { get; private set; } = -1;

        public event Action<string>? OnLog;
        public event Action<int>? OnPaused;
        public event Action? OnStopped;

        // 擬人化設定（外部から設定可能）
        public bool HumanizeEnabled { get; set; } = false;
        public int HumanizeLower { get; set; } = 30;
        public int HumanizeUpper { get; set; } = 100;
        private readonly Random _rng = new Random();

        // New: toggle whether to send keys by scan code at runtime
        public bool UseScanCode { get; set; } = false;

        public ScriptService(InputService input)
        {
            _input = input;
        }

        // Helper to log base/actual/offset timing info in a consistent format
        private void LogTiming(string label, int baseMs, int actualMs, int offset)
        {
            try
            {
                if (HumanizeEnabled && offset != 0)
                    OnLog?.Invoke($"{label}: base={baseMs}ms, 実際={actualMs}ms (擬人化オフセット={offset}ms)");
                else
                    OnLog?.Invoke($"{label}: {actualMs}ms");
            }
            catch { }
        }

        public void Pause()
        {
            _pauseEvent.Reset();
            OnLog?.Invoke("Paused");
            // notify UI of pause and current index
            OnPaused?.Invoke(CurrentIndex);
        }

        public void Resume()
        {
            _pauseEvent.Set();
            OnLog?.Invoke("Resumed");
        }

        public void Stop()
        {
            _cts?.Cancel();
            _pauseEvent.Set();
            OnLog?.Invoke("Stop requested");
        }

        public Task StartAsync(List<ScenarioStep> steps, int startIndex = 0)
        {
            if (IsRunning)
                throw new InvalidOperationException("Already running");

            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            IsRunning = true;

            return Task.Run(async () =>
            {
                try
                {
                    // change: schedule actions relative to previous row end; each action gets its own humanized offset
                    var runTimer = System.Diagnostics.Stopwatch.StartNew();
                    long prevRowEnd = 0; // milliseconds since run start
                    for (int i = startIndex; i < steps.Count; i++)
                    {
                        token.ThrowIfCancellationRequested();
                        _pauseEvent.Wait(token);

                        CurrentIndex = i;
                        var step = steps[i];
                        var actions = (step.Actions != null && step.Actions.Count > 0) ? step.Actions : step.Positions;
                        int actionCount = actions.Count;
                        OnLog?.Invoke($"行{i + 1}: アクション {actionCount} 件スケジュール");

                        int baseDelay = Math.Max(0, step.Delay);
                        int basePd = Math.Max(0, step.PressDuration);

                        // build per-action schedule
                        var schedule = new List<(int idx, string action, long scheduledStartMs, int pdBase, int pdActual, int pdOffset, int delayOffset)>();
                        for (int aidx = 0; aidx < actionCount; aidx++)
                        {
                            var action = actions[aidx]?.Trim() ?? string.Empty;
                            int delayOffset = 0;
                            if (HumanizeEnabled && baseDelay > 0)
                            {
                                try
                                {
                                    int lower = Math.Max(0, HumanizeLower);
                                    int upper = Math.Max(lower, HumanizeUpper);
                                    if (upper > 0)
                                    {
                                        int off;
                                        do { off = _rng.Next(-upper, upper + 1); } while (Math.Abs(off) < lower);
                                        delayOffset = off;
                                    }
                                }
                                catch { }
                            }
                            int actualDelay = Math.Max(0, baseDelay + delayOffset);

                            int pdOffset = 0;
                            if (HumanizeEnabled && basePd > 0)
                            {
                                try
                                {
                                    int lower = Math.Max(0, HumanizeLower);
                                    int upper = Math.Max(lower, HumanizeUpper);
                                    if (upper > 0)
                                    {
                                        int off;
                                        do { off = _rng.Next(-upper, upper + 1); } while (Math.Abs(off) < lower);
                                        pdOffset = off;
                                    }
                                }
                                catch { }
                            }
                            int pdActual = Math.Max(0, basePd + pdOffset);

                            long scheduledStart = prevRowEnd + actualDelay;
                            schedule.Add((aidx, action, scheduledStart, basePd, pdActual, pdOffset, delayOffset));
                        }

                        var ordered = schedule.OrderBy(x => x.scheduledStartMs).ThenBy(x => x.idx).ToList();
                        long rowEndCandidate = prevRowEnd;

                        foreach (var item in ordered)
                        {
                            token.ThrowIfCancellationRequested();
                            _pauseEvent.Wait(token);

                            // wait until scheduled start
                            while (true)
                            {
                                var now = runTimer.ElapsedMilliseconds;
                                var toWait = item.scheduledStartMs - now;
                                if (toWait <= 0) break;
                                var chunk = (int)Math.Min(100, toWait);
                                token.ThrowIfCancellationRequested();
                                _pauseEvent.Wait(token);
                                await Task.Delay(chunk, token).ConfigureAwait(false);
                            }

                            OnLog?.Invoke($"行{i + 1} アクション{item.idx + 1}/{actionCount} 開始 (delay base={baseDelay} offset={item.delayOffset})");
                            OnLog?.Invoke($"アクション解析: '{item.action}'");

                            var parts = item.action.Split(',');
                            if (parts.Length >= 2
                                && double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dx)
                                && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dy))
                            {
                                var xx = (int)Math.Round(dx);
                                var yy = (int)Math.Round(dy);
                                var pl = new PositionList(); pl.Points.Add(new Point(xx, yy));
                                OnLog?.Invoke($"座標実行: {xx},{yy}");
                                LogTiming("押下時間", item.pdBase, item.pdActual, item.pdOffset);
                                _input.ClickMultiple(pl, item.pdActual);
                                rowEndCandidate = Math.Max(rowEndCandidate, runTimer.ElapsedMilliseconds + 0);
                            }
                            else
                            {
                                LogTiming("押下時間", item.pdBase, item.pdActual, item.pdOffset);
                                try
                                {
                                    var handled = false;
                                    try { handled = ExternalSendOverride?.Invoke(item.action, item.pdActual, UseScanCode) ?? false; } catch { handled = false; }
                                    if (!handled)
                                    {
                                        _input.SendByKeyNameWithDuration(item.action, item.pdActual, UseScanCode);
                                    }
                                }
                                catch { SendKeyAction(item.action); }
                                rowEndCandidate = Math.Max(rowEndCandidate, runTimer.ElapsedMilliseconds + 0);
                            }
                        }

                        // row end is when the last action finished
                        prevRowEnd = Math.Max(prevRowEnd, runTimer.ElapsedMilliseconds);
                    }

                    OnLog?.Invoke("実行完了");
                }
                catch (OperationCanceledException)
                {
                    OnLog?.Invoke("実行停止");
                }
                finally
                {
                    IsRunning = false;
                    OnStopped?.Invoke();
                }
            }, token);
        }

        private PositionList ParsePositions(List<string> positions)
        {
            var list = new PositionList();
            foreach (var s in positions.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                var parts = s.Split(',');
                if (parts.Length >= 2
                    && double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dx)
                    && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dy))
                {
                    var x = (int)Math.Round(dx);
                    var y = (int)Math.Round(dy);
                    list.Points.Add(new Point(x, y));
                }
            }
            return list;
        }

        // Example helper used by the service when sending a key action
        private void SendKeyAction(string keySpec)
        {
            try
            {
                try
                {
                    var handled = false;
                    try { handled = ExternalSendOverride?.Invoke(keySpec, 0, UseScanCode) ?? false; } catch { handled = false; }
                    if (!handled)
                    {
                        if (UseScanCode)
                        {
                            // InputService should provide scan-code based sending API
                            _input.SendByScanCode(keySpec);
                            OnLog?.Invoke($"Send by ScanCode: {keySpec}");
                        }
                        else
                        {
                            _input.SendByKeyName(keySpec);
                            OnLog?.Invoke($"Send by KeyName: {keySpec}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    OnLog?.Invoke($"SendKeyAction failed: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                OnLog?.Invoke($"SendKeyAction failed: {ex.Message}");
            }
        }
    }
}
