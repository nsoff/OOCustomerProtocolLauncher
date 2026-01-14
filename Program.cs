using BDOpenOffice;
using System.Diagnostics;
using System.Net;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Linq;
class Program
{
    const string SchemeName = "myoffice"; // 你想要的协议名（myoffice://...）

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool AllocConsole();
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool FreeConsole();

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern int MessageBoxW(System.IntPtr hWnd, string lpText, string lpCaption, uint uType);

    const uint MB_OK = 0x00000000;
    const uint MB_ICONINFORMATION = 0x00000040;
    const uint MB_ICONERROR = 0x00000010;


    const int MaxSavedFiles = 10; // 超过则删除最旧的
    const long MaxTotalBytes = 1024L * 1024L * 500L; // 500 MB 总占用上限

    static int Main(string[] args)
    {
        bool consoleAllocated = false;
        try
        {
            if (ShouldShowConsole(args))
            {
                AllocConsole();
                consoleAllocated = true;
            }

            if (args.Length == 1 && args[0].Equals("--register", StringComparison.OrdinalIgnoreCase))
            {
                if (RegistryHelper.RegisterProtocol(SchemeName))
                    Console.WriteLine($"Protocol '{SchemeName}' registered (HKCU).");
                else
                    Console.WriteLine("Register failed.");
                return 0;
            }
            if (args.Length == 1 && args[0].Equals("--unregister", StringComparison.OrdinalIgnoreCase))
            {
                if (RegistryHelper.UnregisterProtocol(SchemeName))
                    Console.WriteLine($"Protocol '{SchemeName}' unregistered.");
                else
                    Console.WriteLine("Unregister failed.");
                return 0;
            }

            if (args.Length >= 1 && args[0].StartsWith(SchemeName + ":", StringComparison.OrdinalIgnoreCase))
            {
                string raw = args[0].Trim('"');
                HandleUri(raw);
                return 0;
            }

            Console.WriteLine("Usage:");
            Console.WriteLine("  --register     register protocol for current user");
            Console.WriteLine("  --unregister   unregister protocol for current user");
            Console.WriteLine("  <uri>          emulate protocol call, e.g. myoffice://open?file=...");
            return 0;
        }
        finally
        {
        }
    }

    static bool ShouldShowConsole(string[] args)
    {
        if (args.Length == 0) return true;
        if (args.Length == 1)
        {
            var a = args[0];
            if (a.Equals("--register", StringComparison.OrdinalIgnoreCase) || a.Equals("--unregister", StringComparison.OrdinalIgnoreCase))
                return true;
            if (!a.StartsWith(SchemeName + ":", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }
        return false;
    }

    static void HandleUri(string uriStr)
    {
        try
        {
            var uri = new Uri(uriStr);

            if (uri.Host.Equals("open", StringComparison.OrdinalIgnoreCase) ||
                uri.Host.Equals("open/", StringComparison.OrdinalIgnoreCase))
            {
                var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
                var fileParam = query.Get("file");
                if (string.IsNullOrEmpty(fileParam))
                {
                    return;
                }

                var decoded = WebUtility.UrlDecode(fileParam);

                if (decoded.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    var localPath = new Uri(decoded).LocalPath;
                    OpenLocalFileWithAssociatedApp(localPath);
                }
                else if (Uri.TryCreate(decoded, UriKind.Absolute, out var maybeUrl) &&
                         (maybeUrl.Scheme == Uri.UriSchemeHttp || maybeUrl.Scheme == Uri.UriSchemeHttps))
                {
                    var local = DownloadRemoteFile(maybeUrl.AbsoluteUri);
                    if (!string.IsNullOrEmpty(local))
                    {
                        // 清理旧文件
                        TryCleanupDownloads();

                        if (!IsOfficeAvailableForFile(local))
                        {
                            ShowMessage("未检测到 Office 关联应用，文件将使用默认程序打开。若需 Office，请先安装 Office。", "BDOpenOffice", MB_OK | MB_ICONINFORMATION);
                        }

                        OpenLocalFileWithAssociatedApp(local);
                    }
                    else
                    {
                        ShowMessage("无法下载远程文件。", "BDOpenOffice", MB_OK | MB_ICONERROR);
                    }
                }
                else
                {
                    var maybePath = decoded.Trim('"');
                    if (File.Exists(maybePath))
                    {
                        OpenLocalFileWithAssociatedApp(maybePath);
                    }
                }
            }
        }
        catch
        {
        }
    }

    static void ShowMessage(string text, string caption, uint type)
    {
        try
        {
            MessageBoxW(System.IntPtr.Zero, text, caption, type);
        }
        catch { }
    }

    static string EnsureDownloadsDir()
    {
        var dir = Path.Combine(AppContext.BaseDirectory, "downloaded");
        try
        {
            Directory.CreateDirectory(dir);
        }
        catch { }
        return dir;
    }

    static void TryCleanupDownloads()
    {
        try
        {
            var dir = EnsureDownloadsDir();
            var files = new DirectoryInfo(dir).GetFiles().OrderBy(f => f.CreationTimeUtc).ToArray();
            long total = files.Sum(f => f.Length);
            // 删除超过数量限制的最旧文件
            while (files.Length > MaxSavedFiles)
            {
                try { files[0].Delete(); } catch { }
                files = files.Skip(1).ToArray();
            }
            // 删除直到总大小符合限制
            int idx = 0;
            while (total > MaxTotalBytes && idx < files.Length)
            {
                try { total -= files[idx].Length; files[idx].Delete(); } catch { }
                idx++;
            }
        }
        catch { }
    }

    static string DownloadRemoteFile(string url)
    {
        try
        {
            using var client = new HttpClient();
            using var resp = client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();

            var suggestedName = Path.GetFileName(new Uri(url).LocalPath);
            if (string.IsNullOrEmpty(suggestedName))
            {
                suggestedName = "downloaded" + Path.GetExtension(url);
                if (string.IsNullOrEmpty(Path.GetExtension(suggestedName)))
                    suggestedName = "downloaded.bin";
            }

            var dir = EnsureDownloadsDir();
            var filePath = Path.Combine(dir, suggestedName);

            var baseName = Path.GetFileNameWithoutExtension(filePath);
            var ext = Path.GetExtension(filePath);
            int i = 1;
            while (File.Exists(filePath))
            {
                filePath = Path.Combine(dir, $"{baseName}({i++}){ext}");
            }

            using (var fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
            {
                resp.Content.CopyToAsync(fs).GetAwaiter().GetResult();
            }

            try
            {
                var zoneAds = filePath + ":Zone.Identifier";
                if (File.Exists(zoneAds))
                    File.Delete(zoneAds);
            }
            catch { }

            return filePath;
        }
        catch
        {
            return null;
        }
    }

    static bool IsOfficeAvailableForFile(string path)
    {
        try
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            string progId = ext switch
            {
                ".doc" or ".docx" or ".docm" => "Word.Document",
                ".xls" or ".xlsx" or ".xlsm" => "Excel.Sheet",
                ".ppt" or ".pptx" or ".pptm" => "PowerPoint.Show",
                _ => null
            };
            if (progId == null) return false;

            using var key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(progId);
            return key != null;
        }
        catch { return false; }
    }

    static void OpenLocalFileWithAssociatedApp(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch { }
    }
}