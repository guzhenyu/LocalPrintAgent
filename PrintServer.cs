using PdfiumViewer;
using System;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace LocalPrintAgent
{
    public class PrintServer
    {
        private static readonly HttpClient Http = new();
        private static string? s_edgePath;

        private readonly HttpListener _listener = new();
        private readonly AppConfig _config;
        private readonly Action<string> _logger;
        private CancellationTokenSource _cts = new();

        public PrintServer(string prefix, AppConfig config, Action<string> logger)
        {
            _listener.Prefixes.Add(prefix);
            _config = config;
            _logger = logger;
        }

        public void Start()
        {
            _listener.Start();
            _logger($"HTTP listening on: {string.Join(", ", _listener.Prefixes)}");
            _ = LoopAsync(_cts.Token);
        }

        public void Stop()
        {
            try
            {
                _cts.Cancel();
                _listener.Stop();
                _listener.Close();
                _logger("HTTP stopped");
            }
            catch { }
        }

        private async Task LoopAsync(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                HttpListenerContext ctx;
                try
                {
                    ctx = await _listener.GetContextAsync();
                }
                catch
                {
                    break;
                }

                _ = Task.Run(() => HandleAsync(ctx), ct);
            }
        }

        private async Task HandleAsync(HttpListenerContext ctx)
        {
            try
            {
                if (!IsLocalRequest(ctx.Request))
                {
                    WriteJson(ctx, 401, new { ok = false, message = "unauthorized" });
                    return;
                }

                var path = ctx.Request.Url?.AbsolutePath ?? "/";
                if (ctx.Request.HttpMethod == "OPTIONS")
                {
                    WriteEmpty(ctx, 204);
                    return;
                }
                if (ctx.Request.HttpMethod == "GET" && path == "/health")
                {
                    WriteJson(ctx, 200, new { ok = true, message = "alive" });
                    return;
                }

                if (ctx.Request.HttpMethod == "GET" && path == "/printers")
                {
                    var printers = new System.Collections.Generic.List<string>();
                    foreach (string p in PrinterSettings.InstalledPrinters) printers.Add(p);
                    WriteJson(ctx, 200, new { ok = true, printers });
                    return;
                }

                if (ctx.Request.HttpMethod == "POST" && path == "/print")
                {
                    await HandlePrintAsync(ctx);
                    return;
                }

                WriteJson(ctx, 404, new { ok = false, message = "not found" });
            }
            catch (Exception ex)
            {
                _logger("ERROR: " + ex);
                WriteJson(ctx, 500, new { ok = false, message = ex.Message });
            }
        }

        private bool IsLocalRequest(HttpListenerRequest request)
        {
            var remote = request.RemoteEndPoint?.Address;
            if (remote == null) return false;
            if (IPAddress.IsLoopback(remote)) return true;

            remote = NormalizeAddress(remote);
            foreach (var address in GetLocalAddresses())
            {
                if (NormalizeAddress(address).Equals(remote)) return true;
            }

            return false;
        }

        private async Task HandlePrintAsync(HttpListenerContext ctx)
        {
            var req = await ReadJsonAsync<PrintRequest>(ctx.Request);
            ValidateRequest(req);

            var printerName = ResolvePrinterName(req);
            var pdfBytes = await GetPdfBytesAsync(req);

            PrintPdfBytes(req, printerName, pdfBytes);
            WriteJson(ctx, 200, new { ok = true, jobId = req.jobId, message = "printed" });
        }

        private static async Task<T> ReadJsonAsync<T>(HttpListenerRequest request)
        {
            using var sr = new StreamReader(request.InputStream, Encoding.UTF8);
            var body = await sr.ReadToEndAsync();
            var req = JsonSerializer.Deserialize<T>(body, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
            if (req == null) throw new Exception("bad json");
            return req;
        }

        private void ValidateRequest(PrintRequest req)
        {
            if (ShouldUsePdf(req))
            {
                if (string.IsNullOrWhiteSpace(req.pdfUrl) && string.IsNullOrWhiteSpace(req.pdfBase64))
                    throw new Exception("pdfUrl or pdfBase64 required");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(req.htmlBase64))
                    throw new Exception("htmlBase64 required");
            }

            if (req.pageSizeId != 1 && req.pageSizeId != 2)
                throw new Exception("pageSizeId must be 1(A3) or 2(A4)");

            if (!string.IsNullOrWhiteSpace(req.printPageRange)
                && !TryParsePageRange(req.printPageRange, out _, out _))
                throw new Exception("printPageRange invalid");
        }

        private string ResolvePrinterName(PrintRequest req)
        {
            if (!string.IsNullOrWhiteSpace(req.printerName))
                return req.printerName;

            var printerName = req.pageSizeId switch
            {
                1 => _config.A3PrinterName,
                2 => _config.A4PrinterName,
                _ => ""
            };

            if (string.IsNullOrWhiteSpace(printerName))
                throw new Exception(req.pageSizeId == 1 ? "A3 printer not configured" : "A4 printer not configured");

            return printerName;
        }

        private async Task<byte[]> GetPdfBytesAsync(PrintRequest req)
        {
            if (ShouldUsePdf(req))
            {
                if (!string.IsNullOrWhiteSpace(req.pdfUrl))
                    return await DownloadPdfAsync(req.pdfUrl);
                if (!string.IsNullOrWhiteSpace(req.pdfBase64))
                    return Convert.FromBase64String(req.pdfBase64);

                throw new Exception("pdfUrl or pdfBase64 required");
            }

            return await RenderHtmlToPdfAsync(req);
        }

        private static async Task<byte[]> DownloadPdfAsync(string pdfUrl)
        {
            if (Uri.TryCreate(pdfUrl, UriKind.Absolute, out var uri))
            {
                if (uri.IsFile)
                    return await File.ReadAllBytesAsync(uri.LocalPath);

                if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                    return await Http.GetByteArrayAsync(uri);
            }

            if (File.Exists(pdfUrl))
                return await File.ReadAllBytesAsync(pdfUrl);

            throw new Exception("pdfUrl not reachable");
        }

        private void PrintPdfBytes(PrintRequest req, string printerName, byte[] pdfBytes)
        {
            using var ms = new MemoryStream(pdfBytes);
            using var pdf = PdfDocument.Load(ms);
            using var printDoc = pdf.CreatePrintDocument();

            var printerSettings = BuildPrinterSettings(req, printerName);
            var paperSize = FindPaperSize(printerSettings, req.pageSizeId);
            if (paperSize == null)
                throw new Exception(req.pageSizeId == 1 ? "printer does not support A3" : "printer does not support A4");

            printDoc.PrinterSettings = printerSettings;
            printDoc.DefaultPageSettings.PaperSize = paperSize;
            printDoc.DefaultPageSettings.Landscape = !req.isPageOrientationPortrait;
            printDoc.PrintController = new StandardPrintController();

            _logger(
                $"Print job: jobId={req.jobId}, printer={printerName}, " +
                $"isPdf={ShouldUsePdf(req)}, bytes={pdfBytes.Length}, " +
                $"pageSizeId={req.pageSizeId}, portrait={req.isPageOrientationPortrait}, " +
                $"duplexSingleSided={req.isDuplexSingleSided}, range={req.printPageRange ?? ""}"
            );

            printDoc.Print();
        }

        private static PrinterSettings BuildPrinterSettings(PrintRequest req, string printerName)
        {
            var settings = new PrinterSettings
            {
                PrinterName = printerName,
                Copies = 1
            };

            if (!settings.IsValid)
                throw new Exception($"printer not found: {printerName}");

            if (settings.CanDuplex)
                settings.Duplex = req.isDuplexSingleSided ? Duplex.Simplex : Duplex.Vertical;

            if (TryParsePageRange(req.printPageRange, out var from, out var to))
            {
                settings.PrintRange = PrintRange.SomePages;
                settings.FromPage = from;
                settings.ToPage = to;
            }
            else
            {
                settings.PrintRange = PrintRange.AllPages;
            }

            return settings;
        }

        private static PaperSize? FindPaperSize(PrinterSettings settings, int pageSizeId)
        {
            var target = pageSizeId == 1 ? PaperKind.A3 : PaperKind.A4;
            foreach (PaperSize size in settings.PaperSizes)
            {
                if (size.Kind == target) return size;
            }

            var tag = pageSizeId == 1 ? "A3" : "A4";
            foreach (PaperSize size in settings.PaperSizes)
            {
                if (size.PaperName.IndexOf(tag, StringComparison.OrdinalIgnoreCase) >= 0)
                    return size;
            }

            return null;
        }

        private static bool TryParsePageRange(string? value, out int from, out int to)
        {
            from = 0;
            to = 0;

            if (string.IsNullOrWhiteSpace(value)) return false;

            var trimmed = value.Trim();
            var dash = trimmed.IndexOf('-', StringComparison.Ordinal);
            if (dash >= 0)
            {
                var left = trimmed.Substring(0, dash).Trim();
                var right = trimmed.Substring(dash + 1).Trim();
                if (!int.TryParse(left, out from)) return false;
                if (!int.TryParse(right, out to)) return false;
            }
            else
            {
                if (!int.TryParse(trimmed, out from)) return false;
                to = from;
            }

            if (from <= 0 || to <= 0) return false;
            if (to < from) return false;
            return true;
        }

        private static bool ShouldUsePdf(PrintRequest req)
        {
            if (req.isPdf) return true;
            if (!string.IsNullOrWhiteSpace(req.pdfUrl)) return true;
            if (!string.IsNullOrWhiteSpace(req.pdfBase64)) return true;
            return false;
        }

        private async Task<byte[]> RenderHtmlToPdfAsync(PrintRequest req)
        {
            var rawHtml = DecodeHtml(req.htmlBase64 ?? "");
            if (string.IsNullOrWhiteSpace(rawHtml))
                throw new Exception("htmlBase64 required");

            var html = BuildHtmlDocument(rawHtml, req.pageSizeId, req.isPageOrientationPortrait);
            var tmpDir = Path.Combine(Path.GetTempPath(), "LocalPrintAgent");
            Directory.CreateDirectory(tmpDir);

            var htmlPath = Path.Combine(tmpDir, $"print_{Guid.NewGuid():N}.html");
            var pdfPath = Path.Combine(tmpDir, $"print_{Guid.NewGuid():N}.pdf");

            File.WriteAllText(htmlPath, html, Encoding.UTF8);

            try
            {
                var edgePath = GetEdgePath();
                if (string.IsNullOrWhiteSpace(edgePath))
                    throw new Exception("msedge not found for html printing");

                var psi = new ProcessStartInfo
                {
                    FileName = edgePath,
                    Arguments =
                        "--headless --disable-gpu --no-first-run --no-default-browser-check --print-to-pdf-no-header " +
                        $"--print-to-pdf={Quote(pdfPath)} {Quote(new Uri(htmlPath).AbsoluteUri)}",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = false,
                    RedirectStandardError = true
                };

                using var proc = Process.Start(psi);
                if (proc == null) throw new Exception("failed to start msedge");

                var timeoutMs = _config.HtmlToPdfTimeoutMs;
                if (timeoutMs <= 0) timeoutMs = 30000;

                var stderrTask = proc.StandardError.ReadToEndAsync();
                var waitTask = proc.WaitForExitAsync();
                var exited = await Task.WhenAny(waitTask, Task.Delay(timeoutMs)) == waitTask;
                if (!exited)
                {
                    try { proc.Kill(true); } catch { }
                    var timeoutErr = "";
                    try
                    {
                        if (await Task.WhenAny(stderrTask, Task.Delay(2000)) == stderrTask)
                            timeoutErr = stderrTask.Result.Trim();
                    }
                    catch { }
                    if (!string.IsNullOrWhiteSpace(timeoutErr))
                        throw new Exception($"html to pdf timeout: {timeoutErr}");
                    throw new Exception("html to pdf timeout");
                }

                await waitTask;
                var stderr = (await stderrTask).Trim();
                if (!File.Exists(pdfPath))
                {
                    throw new Exception(string.IsNullOrWhiteSpace(stderr) ? "html to pdf failed" : stderr);
                }

                return await File.ReadAllBytesAsync(pdfPath);
            }
            finally
            {
                TryDelete(htmlPath);
                TryDelete(pdfPath);
            }
        }

        private static string DecodeHtml(string htmlBase64)
        {
            if (string.IsNullOrWhiteSpace(htmlBase64)) return "";

            try
            {
                var bytes = Convert.FromBase64String(htmlBase64);
                return Encoding.UTF8.GetString(bytes);
            }
            catch
            {
                return htmlBase64;
            }
        }

        private static string BuildHtmlDocument(string html, int pageSizeId, bool isPortrait)
        {
            var size = GetPageSizeMm(pageSizeId, isPortrait);
            var css = $"@page {{ size: {size.widthMm}mm {size.heightMm}mm; margin: 0; }}";
            var styleTag = $"<style>{css}</style>";

            var source = html;
            var headIndex = source.IndexOf("<head", StringComparison.OrdinalIgnoreCase);
            if (headIndex >= 0)
            {
                var close = source.IndexOf(">", headIndex, StringComparison.OrdinalIgnoreCase);
                if (close >= 0)
                {
                    return html.Insert(close + 1, styleTag);
                }
            }

            var htmlIndex = source.IndexOf("<html", StringComparison.OrdinalIgnoreCase);
            if (htmlIndex >= 0)
            {
                var close = source.IndexOf(">", htmlIndex, StringComparison.OrdinalIgnoreCase);
                if (close >= 0)
                {
                    return html.Insert(close + 1, "<head>" + styleTag + "</head>");
                }
                return styleTag + html;
            }

            return "<!doctype html><html><head><meta charset=\"utf-8\">" +
                   styleTag +
                   "</head><body>" + html + "</body></html>";
        }

        private static (int widthMm, int heightMm) GetPageSizeMm(int pageSizeId, bool isPortrait)
        {
            var width = pageSizeId == 1 ? 297 : 210;
            var height = pageSizeId == 1 ? 420 : 297;
            return isPortrait ? (width, height) : (height, width);
        }

        private static string? GetEdgePath()
        {
            if (!string.IsNullOrWhiteSpace(s_edgePath)) return s_edgePath;

            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            var candidates = new[]
            {
                Path.Combine(programFilesX86, "Microsoft", "Edge", "Application", "msedge.exe"),
                Path.Combine(programFiles, "Microsoft", "Edge", "Application", "msedge.exe")
            };

            foreach (var candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    s_edgePath = candidate;
                    return s_edgePath;
                }
            }

            return null;
        }

        private static string Quote(string value) => "\"" + value + "\"";

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        private void WriteJson(HttpListenerContext ctx, int code, object obj)
        {
            var json = JsonSerializer.Serialize(obj);
            var bytes = Encoding.UTF8.GetBytes(json);

            ctx.Response.StatusCode = code;
            ctx.Response.ContentType = "application/json; charset=utf-8";
            AddCorsHeaders(ctx.Response);
            ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
            ctx.Response.Close();
        }

        private void WriteEmpty(HttpListenerContext ctx, int code)
        {
            ctx.Response.StatusCode = code;
            AddCorsHeaders(ctx.Response);
            ctx.Response.Close();
        }

        private static void AddCorsHeaders(HttpListenerResponse response)
        {
            response.Headers["Access-Control-Allow-Origin"] = "*";
            response.Headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS";
            response.Headers["Access-Control-Allow-Headers"] = "Content-Type";
            response.Headers["Access-Control-Max-Age"] = "86400";
        }

        private static IPAddress NormalizeAddress(IPAddress address)
        {
            return address.IsIPv4MappedToIPv6 ? address.MapToIPv4() : address;
        }

        private static IPAddress[] GetLocalAddresses()
        {
            try
            {
                return Dns.GetHostAddresses(Dns.GetHostName());
            }
            catch
            {
                return Array.Empty<IPAddress>();
            }
        }
    }
}
