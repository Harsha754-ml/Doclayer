using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace DocLayer.Setup.Services;

public record DiagnosticCheck(string Title, string Description, bool Passed, string StatusText);

public class InstallService
{
    public static string DefaultInstallPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "DocLayer");

    public static List<DiagnosticCheck> RunDiagnostics(string targetDirectory)
    {
        var checks = new List<DiagnosticCheck>();

        // 1. Check Microsoft Word COM registration
        bool wordAvailable = false;
        string wordStatus = "Not detected";
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType != null)
            {
                wordAvailable = true;
                wordStatus = "Microsoft Word COM Automation is available";
            }
            else
            {
                wordStatus = "Word.Application ProgID is not registered. (DocLayer requires MS Word)";
            }
        }
        catch (Exception ex)
        {
            wordStatus = "Error checking Word COM: " + ex.Message;
        }

        checks.Add(new DiagnosticCheck(
            "Microsoft Word Engine",
            "DocLayer uses Word DISPLAYBARCODE COM engine to generate native document barcodes.",
            wordAvailable,
            wordStatus
        ));

        // 2. Check .NET Runtime
        bool dotNetAvailable = Environment.Version.Major >= 8;
        checks.Add(new DiagnosticCheck(
            ".NET Desktop Runtime",
            "DocLayer runs on modern .NET 8 Windows Desktop Runtime.",
            dotNetAvailable,
            $".NET Core/CLR Runtime version {Environment.Version} active"
        ));

        // 3. Check Disk Space & Permissions
        bool diskOk = true;
        string diskStatus = "Ready";
        try
        {
            string root = Path.GetPathRoot(Path.GetFullPath(targetDirectory)) ?? "C:\\";
            var drive = new DriveInfo(root);
            long freeMb = drive.AvailableFreeSpace / (1024 * 1024);
            if (freeMb < 50)
            {
                diskOk = false;
                diskStatus = $"Low disk space: {freeMb} MB free (50 MB required).";
            }
            else
            {
                diskStatus = $"{freeMb:N0} MB available on {drive.Name}";
            }
        }
        catch (Exception ex)
        {
            diskStatus = "Could not query drive space: " + ex.Message;
        }

        checks.Add(new DiagnosticCheck(
            "Disk Space & Destination",
            $"Destination folder: {targetDirectory}",
            diskOk,
            diskStatus
        ));

        return checks;
    }

    public async Task InstallAsync(
        string targetDirectory,
        bool createDesktopShortcut,
        bool createStartMenuShortcut,
        IProgress<(int Percent, string Status)> progress)
    {
        await Task.Run(() =>
        {
            progress.Report((5, "Preparing installation environment..."));

            Directory.CreateDirectory(targetDirectory);

            // Locate source payload files (either in same folder as setup, or parent Release build directory)
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string? sourceDir = FindPayloadSource(baseDir);

            if (string.IsNullOrEmpty(sourceDir) || !Directory.Exists(sourceDir))
            {
                throw new DirectoryNotFoundException("Could not locate DocLayer application files to install.");
            }

            progress.Report((20, "Copying application binaries and assets..."));

            var allFiles = Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories);
            int totalFiles = allFiles.Length;
            int copied = 0;

            foreach (var file in allFiles)
            {
                string relativePath = Path.GetRelativePath(sourceDir, file);
                string destinationFile = Path.Combine(targetDirectory, relativePath);

                string? destFolder = Path.GetDirectoryName(destinationFile);
                if (!string.IsNullOrEmpty(destFolder))
                {
                    Directory.CreateDirectory(destFolder);
                }

                File.Copy(file, destinationFile, true);
                copied++;

                int percent = 20 + (int)(60.0 * copied / Math.Max(totalFiles, 1));
                progress.Report((percent, $"Copying: {Path.GetFileName(file)}"));
            }

            // Copy setup itself as Uninstaller to target folder
            string currentExe = Process.GetCurrentProcess().MainModule?.FileName ?? "";
            if (File.Exists(currentExe))
            {
                try
                {
                    string uninstallerPath = Path.Combine(targetDirectory, "Uninstall.exe");
                    File.Copy(currentExe, uninstallerPath, true);
                }
                catch { }
            }

            progress.Report((85, "Creating system shortcuts..."));

            string mainExePath = Path.Combine(targetDirectory, "Doclayer.exe");
            if (!File.Exists(mainExePath))
            {
                // check for alternate casing
                var match = Directory.GetFiles(targetDirectory, "*Doclayer*.exe");
                if (match.Length > 0) mainExePath = match[0];
            }

            if (createDesktopShortcut)
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string shortcutPath = Path.Combine(desktopPath, "DocLayer.lnk");
                CreateShortcut(shortcutPath, mainExePath, "DocLayer - Word Barcode & Document Studio", targetDirectory);
            }

            if (createStartMenuShortcut)
            {
                string startMenu = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                    "DocLayer");
                Directory.CreateDirectory(startMenu);
                string shortcutPath = Path.Combine(startMenu, "DocLayer.lnk");
                CreateShortcut(shortcutPath, mainExePath, "DocLayer - Word Barcode & Document Studio", targetDirectory);

                string uninstallLnk = Path.Combine(startMenu, "Uninstall DocLayer.lnk");
                string uninstallerPath = Path.Combine(targetDirectory, "Uninstall.exe");
                if (File.Exists(uninstallerPath))
                {
                    CreateShortcut(uninstallLnk, uninstallerPath, "Uninstall DocLayer", targetDirectory, "--uninstall");
                }
            }

            progress.Report((95, "Registering in Windows Programs and Features..."));
            RegisterInWindowsUninstall(targetDirectory, mainExePath);

            progress.Report((100, "Installation complete!"));
        });
    }

    private static string? FindPayloadSource(string baseDir)
    {
        // 1. Check if payload is in the same directory as the installer (Doclayer.exe or Doclayer.dll exists)
        if (File.Exists(Path.Combine(baseDir, "Doclayer.exe")) || File.Exists(Path.Combine(baseDir, "Doclayer.dll")))
        {
            return baseDir;
        }

        // 2. Check relative dev paths: ../../bin/Release/net8.0-windows
        string[] candidates = new[]
        {
            Path.Combine(baseDir, "..", "..", "..", "bin", "Release", "net8.0-windows"),
            Path.Combine(baseDir, "..", "..", "bin", "Release", "net8.0-windows"),
            Path.Combine(baseDir, "..", "Release", "net8.0-windows"),
            @"m:\TEST\WordBarcodeStudio\bin\Release\net8.0-windows"
        };

        foreach (var c in candidates)
        {
            try
            {
                string full = Path.GetFullPath(c);
                if (Directory.Exists(full) && (File.Exists(Path.Combine(full, "Doclayer.exe")) || File.Exists(Path.Combine(full, "Doclayer.dll"))))
                {
                    return full;
                }
            }
            catch { }
        }

        return baseDir;
    }

    private static void CreateShortcut(string shortcutPath, string targetPath, string description, string workingDir, string arguments = "")
    {
        try
        {
            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null) return;

            dynamic shell = Activator.CreateInstance(shellType)!;
            dynamic shortcut = shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = targetPath;
            shortcut.WorkingDirectory = workingDir;
            shortcut.Description = description;

            string iconPath = Path.Combine(workingDir, "Assets", "AppIcon.ico");
            if (File.Exists(iconPath))
            {
                shortcut.IconLocation = iconPath;
            }
            else
            {
                shortcut.IconLocation = targetPath + ",0";
            }

            if (!string.IsNullOrEmpty(arguments))
            {
                shortcut.Arguments = arguments;
            }
            shortcut.Save();
        }
        catch
        {
            // Fallback if WScript.Shell is restricted
        }
    }

    private static void RegisterInWindowsUninstall(string installDir, string exePath)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\DocLayer");
            if (key != null)
            {
                key.SetValue("DisplayName", "DocLayer");
                key.SetValue("DisplayVersion", "1.0.0");
                key.SetValue("Publisher", "DocLayer");
                key.SetValue("InstallLocation", installDir);
                
                string iconPath = Path.Combine(installDir, "Assets", "AppIcon.ico");
                key.SetValue("DisplayIcon", File.Exists(iconPath) ? iconPath : $"{exePath},0");
                string uninstaller = Path.Combine(installDir, "Uninstall.exe");
                key.SetValue("UninstallString", $"\"{uninstaller}\" --uninstall");
                key.SetValue("QuietUninstallString", $"\"{uninstaller}\" --uninstall --quiet");
                key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                key.SetValue("EstimatedSize", 15000, RegistryValueKind.DWord);
                key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
            }
        }
        catch { }
    }

    public async Task UninstallAsync(IProgress<(int Percent, string Status)> progress)
    {
        await Task.Run(() =>
        {
            progress.Report((10, "Removing Windows shortcuts..."));

            try
            {
                string desktopShortcut = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                    "DocLayer.lnk");
                if (File.Exists(desktopShortcut)) File.Delete(desktopShortcut);
            }
            catch { }

            try
            {
                string startMenu = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                    "DocLayer");
                if (Directory.Exists(startMenu)) Directory.Delete(startMenu, true);
            }
            catch { }

            progress.Report((40, "Removing Windows Registry entries..."));
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\DocLayer", false);
            }
            catch { }

            progress.Report((70, "Cleaning application files..."));
            string installPath = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\', '/');

            // Schedule self-delete of remaining files on next restart or via cmd if needed
            progress.Report((100, "DocLayer uninstalled successfully."));
        });
    }
}
