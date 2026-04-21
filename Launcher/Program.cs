using System.Diagnostics;
using System.Text;
using System.Windows.Forms;

static class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            var appDir = AppContext.BaseDirectory;
            Directory.SetCurrentDirectory(appDir);

            var configPath = Path.Combine(appDir, "app_config.json");
            var templatePath = Path.Combine(appDir, "app_config.template.json");
            if (!File.Exists(configPath) && File.Exists(templatePath))
            {
                File.Copy(templatePath, configPath, overwrite: false);
            }

            var pythonExe = ResolvePythonExecutable();
            if (pythonExe is null)
            {
                MessageBox.Show(
                    "未找到 Python 解释器，无法启动软件。",
                    "启动错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            var pythonwExe = ResolvePythonwExecutable(pythonExe);
            var scriptPath = Path.Combine(appDir, "tr_raman_ui.py");
            if (!File.Exists(scriptPath))
            {
                MessageBox.Show(
                    "未找到 tr_raman_ui.py。",
                    "启动错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            if (!EnsurePythonPackages(pythonExe, appDir))
            {
                MessageBox.Show(
                    "缺少所需 Python 包，且自动安装失败。",
                    "启动错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = pythonwExe,
                Arguments = $"\"{scriptPath}\"",
                WorkingDirectory = appDir,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.ToString(),
                "启动错误",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    static string? ResolvePythonExecutable()
    {
        var candidates = new[]
        {
            @"C:\Users\adimn\AppData\Local\Programs\Python\Python313\python.exe",
            "python"
        };

        foreach (var candidate in candidates)
        {
            if (candidate.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                if (File.Exists(candidate))
                    return candidate;
            }
            else if (CanRun(candidate, "--version"))
            {
                return candidate;
            }
        }
        return null;
    }

    static string ResolvePythonwExecutable(string pythonExe)
    {
        if (pythonExe.EndsWith("python.exe", StringComparison.OrdinalIgnoreCase))
        {
            var pythonw = pythonExe[..^10] + "pythonw.exe";
            if (File.Exists(pythonw))
                return pythonw;
        }
        return "pythonw";
    }

    static bool EnsurePythonPackages(string pythonExe, string workingDirectory)
    {
        if (CanRun(pythonExe, "-c \"import pyvisa, pyvisa_py\"", workingDirectory))
            return true;

        return CanRun(
            pythonExe,
            "-m pip install pyvisa pyvisa-py psutil zeroconf",
            workingDirectory
        );
    }

    static bool CanRun(string fileName, string arguments, string? workingDirectory = null)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                WorkingDirectory = workingDirectory ?? Environment.CurrentDirectory,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };
            using var process = Process.Start(psi);
            if (process is null)
                return false;
            process.WaitForExit();
            return process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }
}
