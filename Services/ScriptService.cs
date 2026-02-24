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

        public ScriptService(InputService input)
        {
            _input = input;
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
                            var parts = s.Split(',');
                            if (parts.Length >= 2 && int.TryParse(parts[0].Trim(), out var xx) && int.TryParse(parts[1].Trim(), out var yy))
                            {
                                var pl = new PositionList();
                                pl.Points.Add(new Point(xx, yy));
                                _input.ClickMultiple(pl);
                                // small pause between actions
                                await Task.Delay(30, token).ConfigureAwait(false);
                            }
                            else
                            {
                                _input.SendKey(s);
                                await Task.Delay(60, token).ConfigureAwait(false);
                            }
                        }

                        // delay (with optional humanization jitter)
                        var delay = Math.Max(0, step.Delay);
                        if (HumanizeEnabled && delay > 0)
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
                                    delay = Math.Max(0, delay + offset);
                                }
                            }
                            catch { }
                        }
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
                if (parts.Length >= 2 && int.TryParse(parts[0].Trim(), out var x) && int.TryParse(parts[1].Trim(), out var y))
                {
                    list.Points.Add(new Point(x, y));
                }
            }
            return list;
        }
    }
}
