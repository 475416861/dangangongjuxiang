using System;
using System.IO;

namespace MultiToolWin
{
    public static class AppConfig
    {
        private static readonly string ConfigDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MultiToolWin");
        private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.txt");

        public static string OutputRoot { get; set; } = "";
        public static bool LogToFile { get; set; } = false;

        public static void Load()
        {
            try
            {
                if (!File.Exists(ConfigPath)) return;
                foreach (var line in File.ReadAllLines(ConfigPath))
                {
                    if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#")) continue;
                    var i = line.IndexOf('=');
                    if (i < 0) continue;
                    var key = line.Substring(0, i).Trim();
                    var val = line.Substring(i + 1).Trim();
                    if (key.Equals("OutputRoot", StringComparison.OrdinalIgnoreCase)) OutputRoot = val;
                    else if (key.Equals("LogToFile", StringComparison.OrdinalIgnoreCase)) LogToFile = val.Equals("true", StringComparison.OrdinalIgnoreCase) || val == "1";
                }
            }
            catch { /* 忽略读取失败 */ }
        }

        public static void Save()
        {
            try
            {
                Directory.CreateDirectory(ConfigDir);
                File.WriteAllLines(ConfigPath, new[]
                {
                    $"OutputRoot={OutputRoot ?? ""}",
                    $"LogToFile={(LogToFile ? "true" : "false")}",
                });
            }
            catch { /* 忽略写入失败 */ }
        }
    }
}
