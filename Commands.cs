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
        public static Task UpdateFromFolder (string sourcePath, string destinationPath) {
            sourcePath = Path.GetFullPath(sourcePath);
            destinationPath = Path.GetFullPath(destinationPath);

            return Task.Run(() => {
                int i = 1;

                var allFiles = Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\ext\\ts\\"));

                foreach (var f in allFiles) {
                    var localPath = Path.GetFullPath(f).Replace(sourcePath, "").Substring(1);
                    var destFile = Path.Combine(destinationPath, localPath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destFile));
                    File.Copy(f, destFile, true);
                    if (i++ % 3 == 0)
                        Console.WriteLine(localPath);
                    else
                        Console.Write(localPath + " ");
                }

                SetupDesktopIni(destinationPath);
            });
        }

        public static async Task ExtractZipFile (string zipFile, string destinationPath) {
            destinationPath = Path.GetFullPath(destinationPath);

            int i = 1;

            using (var s = File.OpenRead(zipFile))
            using (var zf = new ZipArchive(s, ZipArchiveMode.Read, true, Encoding.UTF8))
            foreach (var entry in zf.Entries) {
                if (entry.FullName.EndsWith("\\") || entry.FullName.EndsWith("/"))
                    continue;

                var destFilename = Path.Combine(destinationPath, entry.FullName);
                Directory.CreateDirectory(Path.GetDirectoryName(destFilename));

                using (var src = entry.Open())
                using (var dst = File.Open(destFilename, FileMode.Create))
                    await src.CopyToAsync(dst);

                File.SetLastWriteTimeUtc(destFilename, entry.LastWriteTime.UtcDateTime);
                if (i++ % 3 == 0)
                    Console.WriteLine(entry.FullName);
                else
                    Console.Write(entry.FullName + " ");
            }

            SetupDesktopIni(destinationPath);
        }

        private static void SetupDesktopIni (string directory) {
            var desktopIni = Path.Combine(directory, "desktop.ini");
            if (File.Exists(desktopIni))
            try {
                (new DirectoryInfo(directory)).Attributes = FileAttributes.System;
                (new FileInfo(desktopIni)).Attributes = FileAttributes.System | FileAttributes.Hidden;
            } catch (Exception exc) {
                Console.Error.WriteLine(exc);
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

            Directory.CreateDirectory(MiscPath);

            var zipPath = Path.Combine(MiscPath, Path.GetFileName(resp.ResponseUri.LocalPath));
            using (var src = resp.GetResponseStream())
            using (var dst = File.Open(zipPath + ".tmp", FileMode.Create))
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

                var newVersionDirectory = Path.Combine(DataPath, "Installer Update");
                Directory.CreateDirectory(newVersionDirectory);

                Console.WriteLine($"Extracting {result.ZipPath} to {newVersionDirectory} ...");
                await ExtractZipFile(result.ZipPath, newVersionDirectory);
                Console.WriteLine($"done.");

                var psi = new ProcessStartInfo(
                    "cmd", 
                    "/C \"timeout /T 5 && echo Updating Viramate Installer... && taskkill /f /im viramate.exe & " +
                    $"copy /Y \"{Path.Combine(newVersionDirectory, "*")}\" \"{ExecutableDirectory}\""
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
            Directory.CreateDirectory(DataPath);

            if (File.Exists(Path.Combine(DataPath, "manifest.json"))) {
                Console.WriteLine("Detected old installation. Removing manifest. You'll need to re-install!");
                onlyIfModified = false;
                File.Delete(Path.Combine(DataPath, "manifest.json"));
            }

            if (installFromDisk.GetValueOrDefault(InstallFromDisk)) {
                Console.WriteLine($"Copying from {DiskSourcePath} to {ExtensionInstallPath} ...");
                await UpdateFromFolder(DiskSourcePath, ExtensionInstallPath);
                Console.WriteLine("done.");

                var managerPath = Path.Combine(ExecutableDirectory, "..", "chromeapp");
                if (Directory.Exists(managerPath)) {
                    Console.WriteLine($"Copying from {managerPath} to {ManagerInstallPath} ...");
                    await UpdateFromFolder(managerPath, ManagerInstallPath);
                    Console.WriteLine("done.");
                }
            } else {
                DownloadResult result;

                Console.Write($"Downloading {ManagerSourceUrl}... ");
                try {
                    result = await DownloadLatest(ManagerSourceUrl);
                    Console.WriteLine($"Extracting {result.ZipPath} to {ManagerInstallPath} ...");
                    await ExtractZipFile(result.ZipPath, ManagerInstallPath);
                    Console.WriteLine($"done.");
                } catch (Exception exc) {
                    Console.Error.WriteLine(exc.Message);
                }

                Console.Write($"Downloading {ExtensionSourceUrl}... ");
                try {
                    result = await DownloadLatest(ExtensionSourceUrl);
                    if (result.WasCached && onlyIfModified)
                        return InstallResult.NotUpdated;

                } catch (Exception exc) {
                    Console.Error.WriteLine(exc.Message);
                    return InstallResult.Failed;
                }

                Console.WriteLine($"Extracting {result.ZipPath} to {ExtensionInstallPath} ...");
                await ExtractZipFile(result.ZipPath, ExtensionInstallPath);
                Console.WriteLine($"done.");
            }

            return InstallResult.Updated;
        }

        static Stream OpenResource (string name) {
            return MyAssembly.GetManifestResourceStream("Viramate." + name.Replace("/", ".").Replace("\\", "."));
        }

        public static async Task InstallExtension () {
            var allowAutoClose = true;

            Console.WriteLine();
            Console.WriteLine($"Viramate Installer v{MyAssembly.GetName().Version}");
            if (Environment.GetCommandLineArgs().Contains("--version"))
                return;

            if (Environment.GetCommandLineArgs().Contains("--update"))
                await AutoUpdateInstaller();

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

                var manifestPath = Path.Combine(MiscPath, "nmh.json");
                Directory.CreateDirectory(MiscPath);
                File.WriteAllText(manifestPath, manifestText);

                Directory.CreateDirectory(Path.Combine(DataPath, "Help"));
                foreach (var n in MyAssembly.GetManifestResourceNames()) {
                    if (!n.EndsWith(".gif") && !n.EndsWith(".png"))
                        continue;

                    var destinationPath = Path.Combine(DataPath, n.Replace("Viramate.", "").Replace("Help.", "Help\\"));
                    using (var src = MyAssembly.GetManifestResourceStream(n))
                    using (var dst = File.Open(destinationPath, FileMode.Create))
                        await src.CopyToAsync(dst);
                }

                const string keyName = @"Software\Google\Chrome\NativeMessagingHosts\com.viramate.installer";
                using (var key = Registry.CurrentUser.CreateSubKey(keyName, true)) {
                    Console.WriteLine($"{keyName}\\@ = {manifestPath}");
                    key.SetValue(null, manifestPath);
                }

                try {
                    WebSocketServer.SetupFirewallRule();
                } catch (Exception exc) {
                    Console.WriteLine("Failed to install firewall rule: {0}", exc);
                    allowAutoClose = false;
                }

                string helpFileText;
                using (var s = new StreamReader(OpenResource("Help/index.html"), Encoding.UTF8))
                    helpFileText = s.ReadToEnd();

                helpFileText = Regex.Replace(
                    helpFileText, 
                    @"\<pre\ id='install_path'>[^<]*\</pre\>", 
                    @"<pre id='install_path'>" + DataPath + "</pre>"
                );

                var helpFilePath = Path.Combine(DataPath, "Help", "index.html");
                File.WriteAllText(helpFilePath, helpFileText);

                Console.WriteLine($"Viramate v{ReadManifestVersion(null)} has been installed.");
                if (!Environment.GetCommandLineArgs().Contains("--nohelp")) {
                    Console.WriteLine("Opening install instructions...");
                    Process.Start(helpFilePath);
                } else if (!Debugger.IsAttached && !IsRunningInsideCmd) {
                    Console.WriteLine("Press enter to exit.");
                    return;
                }

                if (!Environment.GetCommandLineArgs().Contains("--nodir")) {
                    Console.WriteLine("Waiting, then opening install directory...");
                    await Task.Delay(2000);
                    Process.Start(DataPath);
                }
            } else {
                await AutoUpdateInstaller();

                if (!Debugger.IsAttached && !IsRunningInsideCmd) {
                    Console.WriteLine("Failed to install extension. Press enter to exit.");
                    Console.ReadLine();
                } else {
                    Console.WriteLine("Failed to install extension.");
                }
            }
        }
    }

    public enum InstallResult {
        Failed = 0,
        NotUpdated,
        Updated
    }
}
