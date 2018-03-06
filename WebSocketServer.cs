using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Security.Cryptography;
using NetFwTypeLib;

namespace Viramate {
    public class WebSocketServer : IDisposable {
        public HttpListener Listener { get; private set; }
        public bool IsDisposed { get; private set; }

        private readonly object Lock = new object();

        public string Url = "ws://127.0.0.1:8678/";

        public WebSocketServer () {
            Init();
        }

        public async Task<bool> Run () {
            try {
                Listener.Start();
            } catch (Exception exc) {
                Console.Error.WriteLine("Failed to start http server: {0}", exc);
                return false;
            }

            Console.Error.WriteLine();

            try {
                SetupFirewallRule();
            } catch (Exception exc) {
                Console.Error.WriteLine("Failed to set-up firewall rule: {0}", exc);
            }

            HttpListener listener;
            while (!IsDisposed) {
                lock (Lock)
                    listener = Listener;

                var context = await listener.GetContextAsync();
                try {
                    await TryHandleRequest(context);
                } catch (Exception exc) {
                    Console.Error.WriteLine("Unhandled exception: {0}", exc);
                }
            }

            return true;
        }

        public static void SetupFirewallRule () {
            var firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(
                Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));

            try {
                var existingRule = firewallPolicy.Rules.Item("Viramate Installer");

                if (existingRule != null) {
                    if (!existingRule.Enabled)
                        Console.Error.WriteLine("WARNING: Viramate windows firewall rule is currently disabled");

                    return;
                }
            } catch (Exception) {
            }

            var firewallRule = (INetFwRule)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
            firewallRule.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
            firewallRule.Description = "Allows Viramate chrome extension to communicate with installer";
            firewallRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
            firewallRule.Enabled = true;
            firewallRule.InterfaceTypes = "All";
            firewallRule.Protocol = 6; // TCP
            firewallRule.LocalPorts = "8678";
            firewallRule.Name = "Viramate Installer";

            firewallPolicy.Rules.Add(firewallRule);
            Console.Error.WriteLine("Added exception to Windows firewall");
        }

        private bool IsAuthenticated (HttpListenerContext context) {
            if (context == null)
                return false;

            if (context.Request.IsWebSocketRequest)
                return true;

            return false;
        }

        private async Task TryHandleRequest (HttpListenerContext context) {
            context.Response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate");

            if (!IsAuthenticated(context))
                return;

            Console.Error.WriteLine("ws connect {0}", context.Request.RawUrl);

            SocketLoop(context);
        }

        private void Init () {
            Listener = new HttpListener();
            Listener.Prefixes.Add(Url.Replace("ws:", "http:"));
        }

        private static Task Send (WebSocket ws, string text) {
            var buffer = Encoding.UTF8.GetBytes(text);
            return ws.SendAsync(new ArraySegment<byte>(buffer), WebSocketMessageType.Text, true, CancellationToken.None);
        }

        private async Task SocketLoop (HttpListenerContext context) {
            HttpListenerWebSocketContext wsc = null;
            try {
                try {
                    wsc = await context.AcceptWebSocketAsync(null);
                } catch (Exception _) {
                    Console.Error.WriteLine("Accept failed {0}", _);
                    return;
                }
                var ws = wsc.WebSocket;

                await MainLoop(ws);
            } catch (Exception exc) {
                Console.Error.WriteLine(exc);
            } finally {
                if (wsc != null)
                    wsc.WebSocket.Dispose();
            }

            Console.Error.WriteLine("ws disconnect");
        }

        private async Task MainLoop (WebSocket ws) {
            var readBuffer = new ArraySegment<byte>(new byte[10240]);

            WebSocketReceiveResult wsrr;

            while (ws.State == WebSocketState.Open) {
                int readOffset = 0, count = 0;
                try {
                    do {
                        wsrr = await ws.ReceiveAsync(new ArraySegment<byte>(readBuffer.Array, readOffset, readBuffer.Count - readOffset), CancellationToken.None);
                        readOffset += wsrr.Count;
                        count += wsrr.Count;
                    } while (!wsrr.EndOfMessage);
                } catch (Exception exc) {
                    Console.Error.WriteLine(exc);
                    break;
                }

                string json = "";
                JObject obj = null;
                try {
                    json = Encoding.UTF8.GetString(readBuffer.Array, 0, count);
                    obj = (JObject)JsonConvert.DeserializeObject(json);
                } catch (Exception exc) {
                    Console.Error.WriteLine("JSON parse error: {0}", exc);
                    Console.Error.WriteLine("// Failed json blob below //");
                    Console.Error.WriteLine(json);
                    Console.Error.WriteLine("// Failed json blob above //");
                }

                if (obj == null)
                    continue;

                var msgType = (string)obj["type"];
                switch (msgType) {
                    default:
                        Console.Error.WriteLine("Unhandled message {0}", msgType);

                        break;
                }
            }

            Dispose();
        }

        public Task Send<T> (WebSocket ws, T value) {
            var json = JsonConvert.SerializeObject(value);
            var bytes = Encoding.UTF8.GetBytes(json);

            var tcs = new TaskCompletionSource<int>();

            lock (ws) {
                if (IsDisposed) {
                    tcs.SetCanceled();
                    return tcs.Task;
                }

                return ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None);
            }
        }

        public void Dispose () {
            if (IsDisposed)
                return;

            IsDisposed = true;
            Listener.Stop();
            Listener.Close();
        }
    }
}
