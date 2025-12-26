using System;
using System.IO;
using System.Text.Json;

namespace LocalPrintAgent
{
    public class PrintRequest
    {
        public string? jobId { get; set; }
        public bool isPdf { get; set; }
        public string? pdfUrl { get; set; }
        public string? htmlBase64 { get; set; }
        public int pageSizeId { get; set; }
        public bool isPageOrientationPortrait { get; set; } = true;
        public bool isDuplexSingleSided { get; set; } = true;
        public string? printPageRange { get; set; }
        public int? copies { get; set; }

        // Backward compatibility
        public string? pdfBase64 { get; set; }
        public string? printerName { get; set; }
    }

    public class AppConfig
    {
        public string Token { get; set; } = "";
        public string A3PrinterName { get; set; } = "";
        public string A4PrinterName { get; set; } = "";
        public int HtmlToPdfTimeoutMs { get; set; } = 120000;

        public static string GetAppDir()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "LocalPrintAgent"
            );
            return dir;
        }

        public static string GetConfigPath() => Path.Combine(GetAppDir(), "config.json");
        public static string GetLogPath() => Path.Combine(GetAppDir(), "agent.log");

        public static AppConfig LoadOrCreate()
        {
            var dir = GetAppDir();
            Directory.CreateDirectory(dir);

            var path = GetConfigPath();
            if (File.Exists(path))
            {
                var json = File.ReadAllText(path);
                var cfg = JsonSerializer.Deserialize<AppConfig>(json);
                if (cfg != null)
                {
                    if (string.IsNullOrWhiteSpace(cfg.Token))
                    {
                        cfg.Token = Guid.NewGuid().ToString("N");
                        cfg.Save();
                    }
                    return cfg;
                }
            }

            var created = new AppConfig { Token = Guid.NewGuid().ToString("N") };
            created.Save();
            return created;
        }

        public void Save()
        {
            var dir = GetAppDir();
            Directory.CreateDirectory(dir);
            File.WriteAllText(
                GetConfigPath(),
                JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true })
            );
        }

        public static void AppendLog(string line)
        {
            var dir = GetAppDir();
            Directory.CreateDirectory(dir);
            File.AppendAllText(GetLogPath(), $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {line}\r\n");
        }
    }
}
