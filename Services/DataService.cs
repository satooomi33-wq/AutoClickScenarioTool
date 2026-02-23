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

            return Directory.EnumerateFiles(folder, "*.json").Select(Path.GetFileName);
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
    }
}
