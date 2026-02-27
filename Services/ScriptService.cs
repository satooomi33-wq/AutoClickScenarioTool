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
                    for (int i = startIndex; i < steps.Count; i++)
                    {
                        token.ThrowIfCancellationRequested();
                        _pauseEvent.Wait(token);

                        CurrentIndex = i;
                        var step = steps[i];
                        // prepare actions (backward compatible)
                        var actions = (step.Actions != null && step.Actions.Count > 0) ? step.Actions : step.Positions;
                        OnLog?.Invoke($"行{i + 1}: アクション {actions.Count} 件実行");

                        // execute each action: coordinate or key
                        foreach (var a in actions.Where(x => !string.IsNullOrWhiteSpace(x)))
                        {
                            var s = a.Trim();
                            // debug: log raw action for diagnostics
                            OnLog?.Invoke($"アクション解析: '{s}'");
                            var parts = s.Split(',');
                            // require two-axis coordinates. allow decimals and negative; normalize by rounding to int pixels
                            if (parts.Length >= 2
                                && double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dx)
                                && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dy))
                            {
                                var xx = (int)Math.Round(dx);
                                var yy = (int)Math.Round(dy);
                                var pl = new PositionList();
                                pl.Points.Add(new Point(xx, yy));
                                // determine press duration and apply humanization if enabled
                                var pdBase = Math.Max(0, step.PressDuration);
                                var pd = pdBase;
                                int pdOffset = 0;
                                if (HumanizeEnabled && pdBase > 0)
                                {
                                    try
                                    {
                                        int lower = Math.Max(0, HumanizeLower);
                                        int upper = Math.Max(lower, HumanizeUpper);
                                        if (upper > 0)
                                        {
                                            int offset;
                                            do
                                            {
                                                offset = _rng.Next(-upper, upper + 1);
                                            } while (Math.Abs(offset) < lower);
                                            pdOffset = offset;
                                            pd = Math.Max(0, pdBase + pdOffset);
                                        }
                                    }
                                    catch { }
                                }
                                // log coordinate being clicked and press duration info using common formatter
                                OnLog?.Invoke($"座標実行: {xx},{yy}");
                                LogTiming("押下時間", pdBase, pd, pdOffset);
                                _input.ClickMultiple(pl, pd);
                                // small pause between actions — allow target window to become foreground
                                await Task.Delay(120, token).ConfigureAwait(false);
                            }
                            else
                            {
                                // treat as key action
                                // determine press duration and apply humanization for key actions as well
                                var pdBaseKey = Math.Max(0, step.PressDuration);
                                var pdKey = pdBaseKey;
                                int pdKeyOffset = 0;
                                if (HumanizeEnabled && pdBaseKey > 0)
                                {
                                    try
                                    {
                                        int lowerK = Math.Max(0, HumanizeLower);
                                        int upperK = Math.Max(lowerK, HumanizeUpper);
                                        if (upperK > 0)
                                        {
                                            int offsetK;
                                            do
                                            {
                                                offsetK = _rng.Next(-upperK, upperK + 1);
                                            } while (Math.Abs(offsetK) < lowerK);
                                            pdKeyOffset = offsetK;
                                            pdKey = Math.Max(0, pdBaseKey + pdKeyOffset);
                                        }
                                    }
                                    catch { }
                                }
                                LogTiming("押下時間", pdBaseKey, pdKey, pdKeyOffset);
                                try
                                {
                                    // attempt to hold key for the press duration when possible
                                    _input.SendByKeyNameWithDuration(s, pdKey, UseScanCode);
                                }
                                catch
                                {
                                    // fallback to immediate send
                                    SendKeyAction(s);
                                }
                                await Task.Delay(60, token).ConfigureAwait(false);
                            }
                        }

                        // delay (with optional humanization jitter)
                        var baseDelay = Math.Max(0, step.Delay);
                        var delay = baseDelay;
                        int humanizeOffset = 0;
                        if (HumanizeEnabled && baseDelay > 0)
                        {
                            try
                            {
                                int lower = Math.Max(0, HumanizeLower);
                                int upper = Math.Max(lower, HumanizeUpper);
                                if (upper > 0)
                                {
                                    int offset;
                                    // ensure magnitude >= lower and <= upper (allow negative/positive)
                                    do
                                    {
                                        offset = _rng.Next(-upper, upper + 1);
                                    } while (Math.Abs(offset) < lower);
                                    humanizeOffset = offset;
                                    delay = Math.Max(0, baseDelay + humanizeOffset);
                                }
                            }
                            catch { }
                        }

                        // Log delay info using common formatter
                        LogTiming("遅延", baseDelay, delay, humanizeOffset);

                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        while (sw.ElapsedMilliseconds < delay)
                        {
                            token.ThrowIfCancellationRequested();
                            _pauseEvent.Wait(token);
                            await Task.Delay(50, token).ConfigureAwait(false);
                        }
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
            catch (Exception ex)
            {
                OnLog?.Invoke($"SendKeyAction failed: {ex.Message}");
            }
        }
    }
}
