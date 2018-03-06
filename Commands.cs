using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using Newtonsoft.Json;

namespace Viramate {
    static partial class Program {
        public static Task UpdateFromFolder (string sourcePath) {
            return Task.Run(() => {
                int i = 1;

                var allFiles = Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\ext\\ts\\"));

                foreach (var f in allFiles) {
                    var localPath = Path.GetFullPath(f).Replace(sourcePath, "").Substring(1);
                    var destinationPath = Path.Combine(InstallPath, localPath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
                    File.Copy(f, destinationPath, true);
                    if (i++ % 3 == 0)
                        Console.WriteLine(localPath);
                    else
                        Console.Write(localPath + " ");
                }
            });
        }

        public static async Task ExtractZipFile (string zipFile, string destinationPath) {
            int i = 1;

            using (var s = File.OpenRead(zipFile))
            using (var zf = new ZipArchive(s, ZipArchiveMode.Read, true, Encoding.UTF8))
            foreach (var entry in zf.Entries) {
                if (entry.FullName.EndsWith("\\") || entry.FullName.EndsWith("/"))
                    continue;

                var destFilename = Path.Combine(destinationPath, entry.FullName);
                Directory.CreateDirectory(Path.GetDirectoryName(destFilename));

                using (var src = entry.Open())
                using (var dst = File.OpenWrite(destFilename))
                    await src.CopyToAsync(dst);

                File.SetLastWriteTimeUtc(destFilename, entry.LastWriteTime.UtcDateTime);
                if (i++ % 3 == 0)
                    Console.WriteLine(entry.FullName);
                else
                    Console.Write(entry.FullName + " ");
            }
        }

        private static HttpWebRequest MakeUpdateWebRequest (string url) {
            var wr = WebRequest.CreateHttp(url);
            wr.CachePolicy = new RequestCachePolicy(
                Environment.GetCommandLineArgs().Contains("--force")
                    ? RequestCacheLevel.Reload
                    : RequestCacheLevel.Revalidate
            );
            return wr;
        }

        public struct DownloadResult {
            public string ZipPath;
            public bool   WasCached;
        }

        public static async Task<DownloadResult> DownloadLatest (string sourceUrl) {
            var wr = MakeUpdateWebRequest(sourceUrl);
            var resp = await wr.GetResponseAsync();
            if (resp.IsFromCache)
                Console.WriteLine("no new update. Using cached update.");

            var zipPath = Path.Combine(InstallPath, Path.GetFileName(resp.ResponseUri.LocalPath));
            using (var src = resp.GetResponseStream())
            using (var dst = File.OpenWrite(zipPath + ".tmp"))
                await src.CopyToAsync(dst);

            if (File.Exists(zipPath))
                File.Delete(zipPath);
            File.Move(zipPath + ".tmp", zipPath);

            if (!resp.IsFromCache)
                Console.WriteLine(" done.");

            return new DownloadResult {
                ZipPath = zipPath,
                WasCached = resp.IsFromCache
            };
        }

        public static async Task<bool> AutoUpdateInstaller () {
            try {
                Console.Write($"Checking for installer update ... ");

                var result = await DownloadLatest(InstallerSourceUrl);
                if (result.WasCached)
                    return false;

                var newVersionDirectory = Path.Combine(InstallPath, "Installer Update");
                Directory.CreateDirectory(newVersionDirectory);

                Console.WriteLine($"Extracting {result.ZipPath} to {newVersionDirectory} ...");
                await ExtractZipFile(result.ZipPath, newVersionDirectory);
                Console.WriteLine($"done.");

                var psi = new ProcessStartInfo(
                    "cmd", 
                    "/C \"timeout /T 5 && echo Updating Viramate Installer... && " +
                    $"copy /Y \"{Path.Combine(newVersionDirectory, "*")}\" \"{ExecutableDirectory}\" && echo Update OK.\""
                ) {
                    CreateNoWindow = false,
                    UseShellExecute = false,
                    WorkingDirectory = ExecutableDirectory
                };

                using (var proc = Process.Start(psi))
                    Console.WriteLine("Installer will be updated momentarily.");

                return true;
            } catch (Exception exc) {
                Console.Error.WriteLine(exc.Message);
                return false;
            }
        }

        public static async Task<InstallResult> InstallExtensionFiles (bool onlyIfModified, bool? installFromDisk) {
            Directory.CreateDirectory(InstallPath);

            if (installFromDisk.GetValueOrDefault(InstallFromDisk)) {
                var sourcePath = Path.GetFullPath(Path.Combine(ExecutableDirectory, "..", "..", "ext"));
                Console.WriteLine($"Copying from {sourcePath} to {InstallPath} ...");
                await UpdateFromFolder(sourcePath);
                Console.WriteLine("done.");
                return InstallResult.Updated;
            } else {
                DownloadResult result;

                Console.Write($"Downloading {ExtensionSourceUrl}... ");
                try {
                    result = await DownloadLatest(ExtensionSourceUrl);
                    if (result.WasCached && onlyIfModified)
                        return InstallResult.NotUpdated;

                } catch (Exception exc) {
                    Console.Error.WriteLine(exc.Message);
                    return InstallResult.Failed;
                }

                Console.WriteLine($"Extracting {result.ZipPath} to {InstallPath} ...");
                await ExtractZipFile(result.ZipPath, InstallPath);
                Console.WriteLine($"done.");

                return InstallResult.Updated;
            }
        }

        static Stream OpenResource (string name) {
            return MyAssembly.GetManifestResourceStream("Viramate." + name.Replace("/", ".").Replace("\\", "."));
        }

        public static async Task InstallExtension () {
            Console.WriteLine();
            Console.WriteLine($"Viramate Installer v{MyAssembly.GetName().Version}");
            if (Environment.GetCommandLineArgs().Contains("--version"))
                return;

            Console.WriteLine("Installing extension. This'll take a moment...");

            if (await InstallExtensionFiles(false, null) != InstallResult.Failed) {
                Console.WriteLine($"Extension id: {ExtensionId}");

                string manifestText;
                using (var s = new StreamReader(OpenResource("nmh.json"), Encoding.UTF8))
                    manifestText = s.ReadToEnd();

                manifestText = manifestText
                    .Replace(
                        "$executable_path$", 
                        ExecutablePath.Replace("\\", "\\\\").Replace("\"", "\\\"")
                    ).Replace(
                        "$extension_id$", ExtensionId
                    );
                var manifestPath = Path.Combine(InstallPath, "nmh.json");
                File.WriteAllText(manifestPath, manifestText);

                const string keyName = @"Software\Google\Chrome\NativeMessagingHosts\com.viramate.installer";
                using (var key = Registry.CurrentUser.CreateSubKey(keyName, true)) {
                    Console.WriteLine($"{keyName}\\@ = {manifestPath}");
                    key.SetValue(null, manifestPath);
                }

                try {
                    WebSocketServer.SetupFirewallRule();
                } catch (Exception exc) {
                    Console.WriteLine("Failed to install firewall rule: {0}", exc);
                }

                Directory.CreateDirectory(Path.Combine(InstallPath, "Help"));
                foreach (var n in MyAssembly.GetManifestResourceNames()) {
                    if (!n.EndsWith(".png"))
                        continue;

                    var destinationPath = Path.Combine(InstallPath, n.Replace("Viramate.", "").Replace("Help.", "Help\\"));
                    using (var src = MyAssembly.GetManifestResourceStream(n))
                    using (var dst = File.OpenWrite(destinationPath))
                        await src.CopyToAsync(dst);
                }

                string helpFileText;
                using (var s = new StreamReader(OpenResource("Help/index.html"), Encoding.UTF8))
                    helpFileText = s.ReadToEnd();

                helpFileText = Regex.Replace(
                    helpFileText, 
                    @"\<pre\ id='extension_path'>[^<]*\</pre\>", 
                    @"<pre id='extension_path'>" + InstallPath + "</pre>"
                );

                var helpFilePath = Path.Combine(InstallPath, "Help", "index.html");
                File.WriteAllText(helpFilePath, helpFileText);

                Console.WriteLine($"Viramate v{ReadManifestVersion(null)} has been installed.");
                if (!Environment.GetCommandLineArgs().Contains("--nohelp")) {
                    Console.WriteLine("Opening install instructions...");
                    Process.Start(helpFilePath);
                }
            } else {
                Console.WriteLine("Failed to install extension.");
            }

            await AutoUpdateInstaller();

            if (!Debugger.IsAttached && !IsRunningInsideCmd) {
                Console.WriteLine("Press enter to exit.");
                Console.ReadLine();
            }
        }
    }

    public enum InstallResult {
        Failed = 0,
        NotUpdated,
        Updated
    }
}
