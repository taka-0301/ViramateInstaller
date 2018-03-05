using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Win32;
using Newtonsoft.Json;

namespace Viramate {
    class Program {
        static bool Running = true;
        static Assembly MyAssembly;

        static void Main (string[] args) {
            MyAssembly = Assembly.GetExecutingAssembly();

            try {
                if ((args.Length == 0) || !args[0].StartsWith("chrome-extension://"))
                    InstallExtension().Wait();
                else
                    MessagingHostMainLoop().Wait();
            } catch (Exception exc) {
                Console.Error.WriteLine("Uncaught: {0}", exc);
                Environment.ExitCode = 1;

                if (Debugger.IsAttached) {
                    Console.WriteLine("Press enter to exit.");
                    Console.ReadLine();
                }
            }
        }

        static string ExtensionId {
            get {
                /*
                var identity = WindowsIdentity.GetCurrent();
                var binLength = identity.User.BinaryLength;
                var bin = new byte[binLength];
                identity.User.GetBinaryForm(bin, 0);
                Array.Reverse(bin);

                var result = new char[32];
                Array.Copy("viramate".ToCharArray(), result, 8);
                for (int i = 0; i < binLength; i++) {
                    int j = i + 8;
                    if (j >= result.Length)
                        break;
                    result[j] = (char)((int)'a' + (bin[i] % 26));
                }
                return new string(result);
                */

                // FIXME: Generating a new extension ID requires manufacturing a public signing key... ugh
                return "fgpokpknehglcioijejfeebigdnbnokj";
            }
        }

        static bool InstallFromDisk {
            get {
                var defaultDebug = (Debugger.IsAttached || ExecutablePath.EndsWith("ViramateInstaller\\bin\\Viramate.exe"));
                var args = Environment.GetCommandLineArgs();
                if (args.Length <= 1)
                    return defaultDebug;

                return (defaultDebug || (args[1] == "--disk")) && (args[1] != "--network");
            }
        }

        const string SourceUrl = "http://luminance.org/vm/ext.zip";

        static string InstallPath {
            get {
                var folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(folder, "Viramate");
            }
        }

        static string ExecutablePath {
            get {
                var cb = MyAssembly.CodeBase;
                var uri = new Uri(cb);
                return uri.LocalPath;
            }
        }

        static string ExecutableDirectory {
            get {
                return Path.GetDirectoryName(ExecutablePath);
            }
        }

        static Task UpdateFromFolder (string sourcePath) {
            return Task.Run(() => {
                var allFiles = Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\ext\\ts\\"));

                foreach (var f in allFiles) {
                    var localPath = Path.GetFullPath(f).Replace(sourcePath, "").Substring(1);
                    var destinationPath = Path.Combine(InstallPath, localPath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
                    File.Copy(f, destinationPath, true);
                }
            });
        }

        static async Task UpdateFromZipFile (string sourcePath) {
            using (var s = File.OpenRead(sourcePath))
            using (var zf = new ZipArchive(s, ZipArchiveMode.Read, true, Encoding.UTF8))
            foreach (var entry in zf.Entries) {
                if (entry.FullName.EndsWith("\\") || entry.FullName.EndsWith("/"))
                    continue;

                var destFilename = Path.Combine(InstallPath, entry.FullName);
                Directory.CreateDirectory(Path.GetDirectoryName(destFilename));

                using (var src = entry.Open())
                using (var dst = File.OpenWrite(destFilename))
                    await src.CopyToAsync(dst);

                File.SetLastWriteTimeUtc(destFilename, entry.LastWriteTime.UtcDateTime);
            }
        }

        static async Task<string> DownloadLatest (string sourceUrl) {
            var wc = new WebClient();
            var zipPath = Path.Combine(InstallPath, "latest.zip");
            await wc.DownloadFileTaskAsync(sourceUrl, zipPath + ".tmp");
            if (File.Exists(zipPath))
                File.Delete(zipPath);
            File.Move(zipPath + ".tmp", zipPath);
            return zipPath;
        }

        static async Task<bool> InstallExtensionFiles () {
            Directory.CreateDirectory(InstallPath);

            if (InstallFromDisk) {
                var sourcePath = Path.GetFullPath(Path.Combine(ExecutableDirectory, "..", "..", "ext"));
                Console.Write($"Copying from {sourcePath} to {InstallPath} ... ");
                await UpdateFromFolder(sourcePath);
                Console.WriteLine("done.");
                return true;
            } else {
                string zipPath = null;

                Console.Write($"Downloading {SourceUrl}... ");
                try {
                    zipPath = await DownloadLatest(SourceUrl);
                    Console.WriteLine("done.");
                } catch (WebException exc) {
                    Console.WriteLine(exc.Message);
                    return false;
                }

                Console.Write($"Extracting {zipPath} to {InstallPath} ... ");
                await UpdateFromZipFile(zipPath);
                Console.WriteLine($"done.");

                return true;
            }
        }

        static Stream OpenResource (string name) {
            return MyAssembly.GetManifestResourceStream("Viramate." + name.Replace("/", ".").Replace("\\", "."));
        }

        static async Task InstallExtension () {
            Console.WriteLine("Installing extension. This'll take a moment...");

            if (await InstallExtensionFiles()) {
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

                Console.WriteLine("Viramate has been downloaded. Opening install instructions...");
                Process.Start(helpFilePath);
            } else {
                Console.WriteLine("Failed to install extension.");
            }
            Console.WriteLine("Press enter to exit.");
            Console.ReadLine();
        }

        class Msg {
            public string type;
        }

        static async Task MessagingHostMainLoop () {
            var stdin = Console.OpenStandardInput();
            var stdout = Console.OpenStandardOutput();
            var logFilePath = Path.Combine(InstallPath, "installer.log");
            using (var log = new StreamWriter(logFilePath, false, Encoding.UTF8)) {
                log.AutoFlush = true;
                await log.WriteLineAsync($"Installer started as native messaging host. Command line: {Environment.CommandLine}");

                while (Running) {
                    var msg = await ReadMessage<Msg>(stdin);
                    if (msg == null) {
                        await Task.Delay(100);
                        continue;
                    }

                    await TryHandleMessage(log, stdout, msg);
                }

                await log.WriteLineAsync($"Exiting.");
            }
        }

        static async Task TryHandleMessage (StreamWriter log, Stream stdout, Msg msg) {
            await log.WriteLineAsync($"Handling message {msg.type}");

            switch (msg.type) {
                case "extension-startup":
                    await WriteMessage(stdout, new { type = "installer-is-working", installFromDisk = InstallFromDisk });
                    Running = false;
                    return;
                case "exit":
                    await WriteMessage(stdout, new { type = "exiting" });
                    Running = false;
                    return;
            }
        }

        static async Task<T> ReadMessage<T> (Stream stream)
            where T : class 
        {
            var lengthBuf = new byte[4];
            if (await stream.ReadAsync(lengthBuf, 0, 4) != 4)
                return null;

            var lengthBytes = BitConverter.ToInt32(lengthBuf, 0);
            var messageBuf = new byte[lengthBytes];
            if (await stream.ReadAsync(messageBuf, 0, lengthBytes) != lengthBytes)
                return null;

            var json = Encoding.UTF8.GetString(messageBuf);
            var result = JsonConvert.DeserializeObject<T>(json);
            return result;
        }

        static async Task WriteMessage<T> (Stream stream, T message)
            where T : class
        {
            var json = JsonConvert.SerializeObject(message);
            var messageByteLength = Encoding.UTF8.GetByteCount(json);
            var messageBuf = new byte[messageByteLength + 4];
            Array.Copy(BitConverter.GetBytes(messageByteLength), messageBuf, 4);
            Encoding.UTF8.GetBytes(json, 0, json.Length, messageBuf, 4);
            await stream.WriteAsync(messageBuf, 0, messageBuf.Length);
            await stream.FlushAsync();
        }
    }
}
