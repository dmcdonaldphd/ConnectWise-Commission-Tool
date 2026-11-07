// Tiny Windows GUI launcher. Node and Chromium are bundled beside this file.
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

internal static class CommissionLauncher {
    [STAThread]
    private static int Main(string[] args) {
        bool check = args.Length == 1 && args[0] == "--self-check";
        try {
            string root = AppDomain.CurrentDomain.BaseDirectory;
            var info = new ProcessStartInfo {
                FileName = Path.Combine(root, "runtime", "node.exe"),
                Arguments = "\"" + Path.Combine(root, "app", "gui.js") + "\"" + (check ? " --self-check" : ""),
                WorkingDirectory = root,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            info.EnvironmentVariables["PLAYWRIGHT_BROWSERS_PATH"] = Path.Combine(root, "browsers");
            var details = new StringBuilder();
            using (var child = new Process { StartInfo = info }) {
                DataReceivedEventHandler collect = delegate(object sender, DataReceivedEventArgs line) {
                    if (line.Data == null) return;
                    lock (details) { if (details.Length < 12000) details.AppendLine(line.Data); }
                };
                child.OutputDataReceived += collect;
                child.ErrorDataReceived += collect;
                child.Start();
                child.BeginOutputReadLine();
                child.BeginErrorReadLine();
                child.WaitForExit();
                if (child.ExitCode != 0 && !check)
                    MessageBox.Show("The Commission Tool could not finish.\n\n" + details,
                        "PHDTech Commission Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return child.ExitCode;
            }
        } catch (Exception error) {
            if (!check) MessageBox.Show("The Commission Tool could not open. Try running the installer again.\n\n" + error.Message,
                "PHDTech Commission Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
