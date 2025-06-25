using System;
using System.IO.Pipes;
using System.Diagnostics;
using System.Windows.Forms;

namespace MailJumpTray
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using var notifier = new TrayNotifier();
            Application.Run();
        }
    }

    public class TrayNotifier : IDisposable
    {
        private readonly NotifyIcon _icon;
        private readonly NamedPipeServerStream _pipe;
        public TrayNotifier()
        {
            _icon = new NotifyIcon
            {
                Text = "MailJump",
                Icon = SystemIcons.Application,
                Visible = true,
                ContextMenuStrip = new ContextMenuStrip()
            };
            _icon.ContextMenuStrip.Items.Add("Exit", null, (s, e) => Application.Exit());
            _pipe = new NamedPipeServerStream("MailJumpPipe", PipeDirection.In);
            _ = WaitForMessageAsync();
        }

        private async Task WaitForMessageAsync()
        {
            while (true)
            {
                await _pipe.WaitForConnectionAsync();
                using var reader = new StreamReader(_pipe);
                var message = await reader.ReadToEndAsync();
                HandleMAPIMessage(message);
                _pipe.Disconnect();
            }
        }

        private void HandleMAPIMessage(string message)
        {
            // message expected in JSON: { "to": "...", "subject": "...", "body": "...", "attachment": "..." }
            try
            {
                var mapi = System.Text.Json.JsonSerializer.Deserialize<MAPIMessage>(message);
                var args = $"mailto:{mapi?.To}?subject={Uri.EscapeDataString(mapi?.Subject ?? string.Empty)}&body={Uri.EscapeDataString(mapi?.Body ?? string.Empty)}";
                Process.Start(new ProcessStartInfo(args) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing MAPI message: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _icon.Visible = false;
            _icon.Dispose();
            _pipe.Dispose();
        }
    }

    public record MAPIMessage(string To, string Subject, string Body, string? Attachment);
}
