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
            if (args.Length == 0)
                InstallExtension();
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
                return Path.GetDirectoryName(uri.LocalPath);
            }
        }

        static void InstallExtension () {
            Console.WriteLine("Installing extension. This'll take a moment...");

            Console.WriteLine($"Extension id: {ExtensionId}");

            const string keyName = @"Software\Google\Chrome\NativeMessagingHosts\viramate";
            using (var key = Registry.CurrentUser.CreateSubKey(keyName, true)) {
                var manifestPath = Path.Combine(ExecutablePath, "nmh.json");

                Console.WriteLine($"{keyName}\\@ = {manifestPath}");
                key.SetValue(null, manifestPath);
            }

            var helpFilePath = Path.Combine(ExecutablePath, "Help", "index.html");
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
