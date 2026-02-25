using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AutoClickScenarioTool.Services
{
    internal static class KeySpecHelper
    {
        private static readonly HashSet<string> AllowedKeywords = new(StringComparer.OrdinalIgnoreCase)
        {
            "ENTER","TAB","ESC","ESCAPE","BACK","BACKSPACE","SPACE",
            "UP","DOWN","LEFT","RIGHT","HOME","END","INSERT","DELETE",
            "PAGEUP","PAGEDOWN","PGUP","PGDN"
        };

        // 保存時に使う修飾子順
        private static readonly string[] ModifierOrder = new[] { "CTRL", "ALT", "SHIFT" };

        // Validate and normalize. Returns (ok, normalizedValue, errorMessage)
        public static (bool ok, string normalized, string? error) ValidateAndNormalize(string raw)
        {
            if (raw is null) return (true, string.Empty, null);

            var s = raw.Trim();
            if (string.IsNullOrEmpty(s)) return (true, string.Empty, null);

            // 禁止トークン "SC" を含む場合は拒否
            if (s.IndexOf("SC", StringComparison.OrdinalIgnoreCase) >= 0)
                return (false, null, "\"SC\" は入力できません");

            // 全角→半角（簡易）：NFKC 正規化で多くを変換
            s = s.Normalize(System.Text.NormalizationForm.FormKC);

            // 座標パターン X,Y （整数）
            var coordMatch = Regex.Match(s, @"^\s*(\d+)\s*,\s*(\d+)\s*$");
            if (coordMatch.Success)
            {
                var x = coordMatch.Groups[1].Value;
                var y = coordMatch.Groups[2].Value;
                return (true, $"{x},{y}", null);
            }

            // キースペック処理: MOD1+MOD2+MAIN
            var parts = s.Split('+').Select(p => p.Trim()).Where(p => p.Length > 0).ToArray();
            if (parts.Length == 0) return (true, string.Empty, null);

            var main = parts.Last();
            var mods = parts.Take(parts.Length - 1).ToArray();

            // 修飾子の検証
            foreach (var m in mods)
            {
                var upm = m.ToUpperInvariant();
                if (upm != "CTRL" && upm != "ALT" && upm != "SHIFT")
                    return (false, null, $"不明な修飾子: {m}");
            }

            // メインキー検証と正規化
            var mainHalf = main; // NFKC済み
            if (mainHalf.Length == 1)
            {
                var ch = mainHalf[0];
                if (char.IsLetter(ch))
                {
                    mainHalf = char.ToUpperInvariant(ch).ToString();
                    // Note: Shift の有無は実行時に判断する設計なのでここでは大文字に統一
                }
                else if (char.IsDigit(ch) || char.IsPunctuation(ch) || char.IsSymbol(ch))
                {
                    // 許可
                }
                else
                {
                    return (false, null, $"無効なキー: {main}");
                }
            }
            else
            {
                if (int.TryParse(mainHalf, out _))
                {
                    // 数字キーとして許可
                }
                else
                {
                    var up = mainHalf.ToUpperInvariant();
                    if (AllowedKeywords.Contains(up) || (up.StartsWith("F") && int.TryParse(up.Substring(1), out int f) && f >= 1 && f <= 24))
                    {
                        mainHalf = up;
                    }
                    else
                    {
                        return (false, null, $"無効なキー名: {main}");
                    }
                }
            }

            // 修飾子を標準順に整列して結合
            var normalizedMods = new List<string>();
            foreach (var mo in ModifierOrder)
            {
                if (mods.Any(m => string.Equals(m, mo, StringComparison.OrdinalIgnoreCase)))
                    normalizedMods.Add(mo);
            }
            var partsOut = normalizedMods.ToList();
            partsOut.Add(mainHalf.ToUpperInvariant());
            var normalized = string.Join("+", partsOut);
            return (true, normalized, null);
        }
    }
}
