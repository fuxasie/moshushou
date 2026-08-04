using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;

namespace moshushou.Input
{
    internal static class DriverInstallationManager
    {
        private const int SuccessRebootRequired = 3010;

        public static bool EnsureVirtualHidInstalled(SearchConfig config)
        {
            if (!string.Equals(
                    config.InputBackend,
                    InputBackendFactory.VirtualHidMode,
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (VirtualHidBackend.IsCompatibleDevicePresent())
            {
                return true;
            }

            MessageBoxResult choice = MessageBox.Show(
                "未检测到 Moshushou Virtual HID 驱动。\n\n是否现在安装？安装过程需要管理员权限，完成后软件会自动重新检测。",
                "安装 Virtual HID 驱动",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.Yes);
            if (choice != MessageBoxResult.Yes)
            {
                DebugLogManager.Log("Input", "用户取消安装 Virtual HID 驱动。");
                return false;
            }

            string baseDirectory = AppContext.BaseDirectory;
            string driverDirectory = Path.Combine(baseDirectory, "Driver");
            string packageDirectory = Path.Combine(driverDirectory, "package");
            string installerPath = Path.Combine(driverDirectory, "Moshushou.DriverInstaller.exe");
            string infPath = Path.Combine(packageDirectory, "MoshushouVirtualHid.inf");
            string certificatePath = Path.Combine(packageDirectory, "MoshushouVirtualHidTest.cer");

            if (!File.Exists(installerPath) || !File.Exists(infPath))
            {
                string message =
                    "驱动安装文件不完整。\n\n" +
                    $"安装器：{installerPath}\n" +
                    $"INF：{infPath}";
                DebugLogManager.Log("Input", message);
                MessageBox.Show(message, "无法安装驱动", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = installerPath,
                    UseShellExecute = true,
                    Verb = "runas",
                    WorkingDirectory = driverDirectory
                };
                startInfo.ArgumentList.Add("install");
                startInfo.ArgumentList.Add("--inf");
                startInfo.ArgumentList.Add(infPath);
                if (File.Exists(certificatePath))
                {
                    startInfo.ArgumentList.Add("--certificate");
                    startInfo.ArgumentList.Add(certificatePath);
                }

                using Process? process = Process.Start(startInfo);
                if (process == null)
                {
                    throw new InvalidOperationException("无法启动驱动安装器。");
                }

                process.WaitForExit();
                if (process.ExitCode != 0 && process.ExitCode != SuccessRebootRequired)
                {
                    string failure = $"Virtual HID 驱动安装失败，退出码：{process.ExitCode}。";
                    DebugLogManager.Log("Input", failure);
                    MessageBox.Show(failure, "驱动安装失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                for (int attempt = 0; attempt < 30; attempt++)
                {
                    if (VirtualHidBackend.IsCompatibleDevicePresent())
                    {
                        config.InputBackend = InputBackendFactory.VirtualHidMode;
                        config.Save();
                        DebugLogManager.Log("Input", "Virtual HID 驱动安装并检测成功。");
                        MessageBox.Show(
                            "Virtual HID 驱动安装成功。",
                            "安装完成",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                        return true;
                    }
                    Thread.Sleep(200);
                }

                string notReady = process.ExitCode == SuccessRebootRequired
                    ? "驱动安装完成，但系统要求重启后才能使用。"
                    : "驱动安装完成，但暂未检测到控制设备。请重新启动软件或电脑。";
                DebugLogManager.Log("Input", notReady);
                MessageBox.Show(notReady, "驱动尚未就绪", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                DebugLogManager.Log("Input", "用户取消了驱动安装的管理员授权。");
                return false;
            }
            catch (Exception ex)
            {
                DebugLogManager.Log("Input", $"驱动安装异常：{ex}");
                MessageBox.Show(
                    $"驱动安装异常：{ex.Message}",
                    "驱动安装失败",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
        }

        public static bool UninstallVirtualHidDriver(
            SearchConfig config,
            Window? owner = null,
            Action? beforeUninstall = null)
        {
            MessageBoxResult choice = MessageBox.Show(
                owner,
                "确定卸载 Moshushou Virtual HID 驱动吗？\n\n" +
                "卸载后键鼠模拟方式将切换为 SendInput，并需要重新启动软件。",
                "卸载 Virtual HID 驱动",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);
            if (choice != MessageBoxResult.Yes)
            {
                return false;
            }

            string baseDirectory = AppContext.BaseDirectory;
            string driverDirectory = Path.Combine(baseDirectory, "Driver");
            string packageDirectory = Path.Combine(driverDirectory, "package");
            string installerPath = Path.Combine(driverDirectory, "Moshushou.DriverInstaller.exe");
            string infPath = Path.Combine(packageDirectory, "MoshushouVirtualHid.inf");

            if (!File.Exists(installerPath) || !File.Exists(infPath))
            {
                string message =
                    "驱动卸载文件不完整。\n\n" +
                    $"安装器：{installerPath}\n" +
                    $"INF：{infPath}";
                DebugLogManager.Log("Input", message);
                MessageBox.Show(owner, message, "无法卸载驱动", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            try
            {
                beforeUninstall?.Invoke();

                var startInfo = new ProcessStartInfo
                {
                    FileName = installerPath,
                    UseShellExecute = true,
                    Verb = "runas",
                    WorkingDirectory = driverDirectory
                };
                startInfo.ArgumentList.Add("uninstall");
                startInfo.ArgumentList.Add("--inf");
                startInfo.ArgumentList.Add(infPath);

                using Process? process = Process.Start(startInfo);
                if (process == null)
                {
                    throw new InvalidOperationException("无法启动驱动卸载程序。");
                }

                process.WaitForExit();
                if (process.ExitCode != 0 && process.ExitCode != SuccessRebootRequired)
                {
                    string failure = $"Virtual HID 驱动卸载失败，退出码：{process.ExitCode}。";
                    DebugLogManager.Log("Input", failure);
                    MessageBox.Show(owner, failure, "驱动卸载失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                config.InputBackend = InputBackendFactory.SendInputMode;
                config.AllowSendInputFallback = true;
                config.Save();

                bool stillAvailable = false;
                for (int attempt = 0; attempt < 20; attempt++)
                {
                    stillAvailable = VirtualHidBackend.IsCompatibleDevicePresent();
                    if (!stillAvailable)
                    {
                        break;
                    }
                    Thread.Sleep(100);
                }

                string resultMessage;
                MessageBoxImage resultIcon;
                if (process.ExitCode == SuccessRebootRequired || stillAvailable)
                {
                    resultMessage = "Virtual HID 驱动卸载操作已完成。请重新启动电脑和软件以完成清理。";
                    resultIcon = MessageBoxImage.Warning;
                }
                else
                {
                    resultMessage = "Virtual HID 驱动已卸载。键鼠模拟方式已切换为 SendInput，请重新启动软件。";
                    resultIcon = MessageBoxImage.Information;
                }

                DebugLogManager.Log("Input", resultMessage);
                MessageBox.Show(owner, resultMessage, "驱动卸载完成", MessageBoxButton.OK, resultIcon);
                return true;
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                DebugLogManager.Log("Input", "用户取消了驱动卸载的管理员授权。");
                return false;
            }
            catch (Exception ex)
            {
                DebugLogManager.Log("Input", $"驱动卸载异常：{ex}");
                MessageBox.Show(
                    owner,
                    $"驱动卸载异常：{ex.Message}",
                    "驱动卸载失败",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
        }
    }
}
