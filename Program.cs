using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Viramate {
    class Program {
        static void Main (string[] args) {
            try {
                if (args.Length == 0)
                    InstallExtension().Wait();
            } catch (Exception exc) {
                Console.Error.WriteLine("Uncaught: {0}", exc);
                Environment.ExitCode = 1;
            }
        }

        static string ExtensionId {
            get {
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
            }
        }

        static bool IsDebugMode {
            get {
                return (Debugger.IsAttached && ExecutablePath.EndsWith("\\bin\\Viramate.exe")) ||
                    (Environment.GetCommandLineArgs()[0] == "--debug");
            }
        }

        static string InstallPath {
            get {
                var folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(folder, "Viramate");
            }
        }

        static string ExecutablePath {
            get {
                var cb = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                var uri = new Uri(cb);
                return uri.LocalPath;
            }
        }

        static string ExecutableDirectory {
            get {
                return Path.GetDirectoryName(ExecutablePath);
            }
        }

        static async Task UpdateExtensionFiles (bool cleanInstall) {
            Directory.CreateDirectory(InstallPath);

            if (IsDebugMode) {
                var sourcePath = Path.GetFullPath(Path.Combine(ExecutableDirectory, "..", "..", "ext"));
                Console.Write($"Copying from {sourcePath} to {InstallPath} ");
                var allFiles = Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories)
                    .Where(f => !f.Contains("\\ext\\ts\\"));
                foreach (var f in allFiles) {
                    var localPath = Path.GetFullPath(f).Replace(sourcePath, "").Substring(1);
                    var destinationPath = Path.Combine(InstallPath, localPath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
                    File.Copy(f, destinationPath, true);
                    Console.Write(".");
                }
                Console.WriteLine();
                Console.WriteLine("done.");
                // Directory.GetFiles(
            }

            return;
        }

        static async Task InstallExtension () {
            if (IsDebugMode)
                Console.WriteLine("// DEBUG MODE");
            Console.WriteLine("Installing extension. This'll take a moment...");

            await UpdateExtensionFiles(true);

            Console.WriteLine($"Extension id: {ExtensionId}");

            var manifestText = File.ReadAllText(Path.Combine(ExecutableDirectory, "nmh.json"));
            manifestText = manifestText
                .Replace(
                    "$executable_path$", 
                    ExecutablePath.Replace("\\", "\\\\").Replace("\"", "\\\"")
                ).Replace(
                    "$extension_id$", ExtensionId
                );
            var manifestPath = Path.Combine(InstallPath, "nmh.json");
            File.WriteAllText(manifestPath, manifestText);

            const string keyName = @"Software\Google\Chrome\NativeMessagingHosts\viramate";
            using (var key = Registry.CurrentUser.CreateSubKey(keyName, true)) {
                Console.WriteLine($"{keyName}\\@ = {manifestPath}");
                key.SetValue(null, manifestPath);
            }

            var helpFilePath = Path.Combine(ExecutableDirectory, "Help", "index.html");
            var text = File.ReadAllText(helpFilePath);
            text = Regex.Replace(text, @"\<pre\ id='extension_path'>[^<]*\</pre\>", @"<pre id='extension_path'>" + InstallPath + "</pre>");
            File.WriteAllText(helpFilePath, text);

            Console.WriteLine("Viramate has been downloaded. Opening install instructions...");
            Process.Start(helpFilePath);
            Console.WriteLine("Press enter to exit.");
            Console.ReadLine();
        }
    }
}
