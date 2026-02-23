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

        public event Action<string>? OnLog;
        public event Action? OnStopped;

        public ScriptService(InputService input)
        {
            _input = input;
        }

        public void Pause()
        {
            _pauseEvent.Reset();
            OnLog?.Invoke("Paused");
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

                        var step = steps[i];
                        OnLog?.Invoke($"行{i + 1}: 座標 {step.Positions.Count} 点クリック");

                        // parse positions
                        var posList = ParsePositions(step.Positions);
                        _input.ClickMultiple(posList);

                        // delay
                        var delay = Math.Max(0, step.Delay);
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
