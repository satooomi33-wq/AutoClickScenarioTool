using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using AutoClickScenarioTool.Models;

namespace AutoClickScenarioTool.Services
{
    public class DataService
    {
        public IEnumerable<string> ListJsonFiles(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                return Enumerable.Empty<string>();

            // Path.GetFileName can return null in some API annotations; filter nulls and cast to non-nullable
            return Directory.EnumerateFiles(folder, "*.json")
                .Select(Path.GetFileName)
                .Where(n => n != null)
                .Select(n => n!);
        }

        public async Task<List<ScenarioStep>> LoadAsync(string filePath)
        {
            if (!File.Exists(filePath))
                return new List<ScenarioStep>();

            var text = await File.ReadAllTextAsync(filePath).ConfigureAwait(false);
            var data = JsonSerializer.Deserialize<List<ScenarioStep>>(text, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
            return data ?? new List<ScenarioStep>();
        }

        public async Task SaveAsync(string filePath, List<ScenarioStep> steps)
        {
            var json = JsonSerializer.Serialize(steps, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(filePath, json).ConfigureAwait(false);
        }

        public async Task<AutoClickScenarioTool.Models.DefaultSettings> LoadDefaultsAsync(string path = "Data/defaults.json")
        {
            if (!File.Exists(path))
                return new AutoClickScenarioTool.Models.DefaultSettings();

            var text = await File.ReadAllTextAsync(path).ConfigureAwait(false);
            try
            {
                var data = JsonSerializer.Deserialize<AutoClickScenarioTool.Models.DefaultSettings>(text, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                // Do not honor persisted HumanizeEnabled value; always use in-code default for initial state
                var result = data ?? new AutoClickScenarioTool.Models.DefaultSettings();
                try { result.HumanizeEnabled = new AutoClickScenarioTool.Models.DefaultSettings().HumanizeEnabled; } catch { }
                return result;
            }
            catch
            {
                return new AutoClickScenarioTool.Models.DefaultSettings();
            }
        }

        public async Task SaveDefaultsAsync(AutoClickScenarioTool.Models.DefaultSettings settings, string path = "Data/defaults.json")
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            // Do not persist the HumanizeEnabled toggle to defaults file by design
            var toSave = new
            {
                Delay = settings.Delay,
                PressDuration = settings.PressDuration,
                HumanizeLower = settings.HumanizeLower,
                HumanizeUpper = settings.HumanizeUpper,
                UseScanCode = settings.UseScanCode
            };
            var json = JsonSerializer.Serialize(toSave, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(path, json).ConfigureAwait(false);
        }
    }
}
