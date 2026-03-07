using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace WindowsSmartTaskbar
{
    public class ProgramItem
    {
        public string Name { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string? Arguments { get; set; }
        public string Category { get; set; } = "All programs";
        public Icon? Icon { get; set; }
        public DateTime AddedDate { get; set; } = DateTime.Now;

        public ProgramItem() { }

        public ProgramItem(string name, string filePath, string? arguments = null)
        {
            Name = name;
            FilePath = filePath;
            Arguments = arguments;
            LoadIcon();
        }

        private void LoadIcon()
        {
            try
            {
                // Försök hämta ikon direkt från filen (fungerar för både .exe och .lnk)
                if (File.Exists(FilePath))
                {
                    Icon = Icon.ExtractAssociatedIcon(FilePath);
                    return;
                }

                // Om originalfilen inte finns, försök med målfilen för genvägar
                if (Path.GetExtension(FilePath).ToLower() == ".lnk")
                {
                    string targetPath = GetShortcutTarget(FilePath);
                    if (File.Exists(targetPath))
                        Icon = Icon.ExtractAssociatedIcon(targetPath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{nameof(ProgramItem)}::{nameof(LoadIcon)}] {ex}");
            }
        }

        private string GetShortcutTarget(string shortcutPath)
        {
            try
            {
                string? target = ReadShortcutTarget(shortcutPath);
                return string.IsNullOrWhiteSpace(target) ? shortcutPath : target;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{nameof(ProgramItem)}::{nameof(GetShortcutTarget)}] {ex}");
                return shortcutPath;
            }
        }

        private string? ReadShortcutTarget(string shortcutPath)
        {
            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                return null;
            }

            object? shell = null;
            object? shortcut = null;
            try
            {
                shell = Activator.CreateInstance(shellType);
                if (shell == null)
                {
                    return null;
                }

                shortcut = shellType.InvokeMember(
                    "CreateShortcut",
                    BindingFlags.InvokeMethod,
                    null,
                    shell,
                    new object[] { shortcutPath });

                if (shortcut == null)
                {
                    return null;
                }

                string? targetPath = shortcut.GetType().InvokeMember(
                    "TargetPath",
                    BindingFlags.GetProperty,
                    null,
                    shortcut,
                    null) as string;

                return string.IsNullOrWhiteSpace(targetPath) ? null : targetPath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{nameof(ProgramItem)}::{nameof(ReadShortcutTarget)}] {ex}");
                return null;
            }
            finally
            {
                if (shortcut != null && Marshal.IsComObject(shortcut))
                {
                    Marshal.FinalReleaseComObject(shortcut);
                }

                if (shell != null && Marshal.IsComObject(shell))
                {
                    Marshal.FinalReleaseComObject(shell);
                }
            }
        }

        public void Start()
        {
            try
            {
                string targetPath = FilePath;
                string targetArgs = Arguments ?? string.Empty;
                
                // Om det är en genväg, försök att starta den direkt via shell
                if (Path.GetExtension(FilePath).ToLower() == ".lnk")
                {
                    // Försök först att starta genvägen direkt
                    try
                    {
                        var startInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = FilePath,
                            UseShellExecute = true
                        };
                        System.Diagnostics.Process.Start(startInfo);
                        return;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[{nameof(ProgramItem)}::{nameof(Start)}] Failed launching shortcut directly: {ex}");
                        targetPath = GetShortcutTarget(FilePath);
                    }
                }

                // Försök att starta programmet
                if (!string.IsNullOrEmpty(targetPath))
                {
                    var startInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = targetPath,
                        Arguments = targetArgs,
                        UseShellExecute = true
                    };
                    System.Diagnostics.Process.Start(startInfo);
                }
                else
                {
                    MessageBox.Show($"Kunde inte hitta målfilen: {targetPath}", "Fel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunde inte starta programmet: {ex.Message}\nSökväg: {FilePath}", "Fel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
