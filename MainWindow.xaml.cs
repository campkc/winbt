using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;

namespace WinBorg
{
    public partial class MainWindow : Window
    {
        // SSH configuration
        private Dictionary<string, string> _sshConfig = new Dictionary<string, string>();

        // Path to borg executable
        private string _borgPath = @"C:\Program Files\BorgBackup\borg.exe";

        public MainWindow()
        {
            // Global exception handler for the UI thread
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                Log($"UNHANDLED EXCEPTION: {ex?.Message ?? "Unknown"}\n{ex?.StackTrace ?? "No stack trace"}");
            };

            try
            {
                Log("MainWindow constructor started");
                InitializeComponent();
                Log("InitializeComponent completed");

                try
                {
                    // Set default values
                    comboCompressionType.SelectedIndex = 1; // Lz4
                    Log("comboCompressionType set to Lz4");
                    comboDeduplication.SelectedIndex = 0; // Small chunks
                    Log("comboDeduplication set to Small chunks");
                    sliderCompressionLevel.Value = 6;
                    Log("sliderCompressionLevel set to 6");

                    UpdateStatus("WinBorg initialized. Ready for operation.");
                    Log("UpdateStatus called with initialization message");
                }
                catch (Exception ex)
                {
                    Log($"Exception setting default values: {ex.Message}\n{ex.StackTrace}");
                }

                try
                {
                    CheckBorgInstallation();
                    Log("CheckBorgInstallation completed");
                }
                catch (Exception ex)
                {
                    Log($"Exception in CheckBorgInstallation (handled): {ex.Message}\n{ex.StackTrace}");
                    // Continue without crashing
                    UpdateStatus($"Warning: Could not verify Borg installation: {ex.Message}", true);
                }

                Log("MainWindow constructor completed");
            }
            catch (Exception ex)
            {
                Log($"Exception in MainWindow constructor: {ex.Message}\n{ex.StackTrace}");
                // Try to show error and continue
                try
                {
                    MessageBox.Show($"Error initializing application: {ex.Message}",
                                   "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch
                {
                    // Last resort - can't even show message box
                    Log("Critical error: Could not show error message box");
                }
            }
        }

        private void CheckBorgInstallation()
        {
            try
            {
                Log("CheckBorgInstallation started");
                if (!File.Exists(_borgPath))
                {
                    UpdateStatus("Borg executable not found at expected location. Please ensure Borg is installed.", true);
                    MessageBox.Show("Borg executable not found at expected location:\n" + _borgPath +
                                    "\n\nPlease ensure Borg is installed correctly.",
                                    "Installation Check", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("Borg executable not found");
                }
                else
                {
                    // Test borg version
                    try
                    {
                        Log("About to check Borg version");
                        List<string> args = new List<string> { "--version" };
                        Log("Executing Borg command with args: " + string.Join(" ", args));

                        var result = new Dictionary<string, object>();

                        // Run on a separate thread to avoid UI freeze during check
                        var task = Task.Run(() => {
                            try
                            {
                                return ExecuteBorgCommand(args);
                            }
                            catch (Exception ex)
                            {
                                Log($"Task exception: {ex.Message}");
                                return new Dictionary<string, object>
                                {
                                    ["exitCode"] = -1,
                                    ["stderr"] = $"Task error: {ex.Message}"
                                };
                            }
                        });

                        // Wait with a timeout
                        if (task.Wait(10000))  // 10-second timeout
                        {
                            result = task.Result;
                        }
                        else
                        {
                            Log("Borg version check timed out");
                            UpdateStatus("Borg version check timed out. Continuing anyway.", true);
                            return;
                        }

                        Log("ExecuteBorgCommand completed for version check");

                        if (result.ContainsKey("exitCode") && (int)result["exitCode"] == 0 && result.ContainsKey("stdout"))
                        {
                            string version = result["stdout"].ToString().Trim();
                            UpdateStatus("Borg found: " + version);
                            Log("Borg found: " + version);
                        }
                        else
                        {
                            string errorMsg = result.ContainsKey("stderr") ? result["stderr"].ToString() : "Unknown error";
                            UpdateStatus("Borg installation check failed: " + errorMsg, true);
                            Log("Borg installation check failed: " + errorMsg);
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus("Error checking Borg: " + ex.Message, true);
                        Log($"Exception in CheckBorgInstallation (inner): {ex.Message}\n{ex.StackTrace}");
                    }
                }

                Log("CheckBorgInstallation about to return");
            }
            catch (Exception ex)
            {
                Log($"Exception in CheckBorgInstallation: {ex.Message}\n{ex.StackTrace}");
                // Don't throw - just log and continue
                UpdateStatus("Error checking Borg installation: " + ex.Message, true);
            }
        }

        #region Repository Type Handling

        private void RepositoryTypeChanged(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("RepositoryTypeChanged started");
                // Toggle visibility of repository path grid
                if (localRepoPathGrid != null)
                {
                    if (radioLocalRepository.IsChecked == true)
                    {
                        localRepoPathGrid.Visibility = Visibility.Visible;
                        Log("localRepoPathGrid set to Visible");
                    }
                    else
                    {
                        localRepoPathGrid.Visibility = Visibility.Collapsed;
                        Log("localRepoPathGrid set to Collapsed");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in RepositoryTypeChanged: {ex.Message}\n{ex.StackTrace}");
                // Don't throw - prevent UI crash
            }
        }

        private void BtnConfigureSSH_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("BtnConfigureSSH_Click started");
                SSHConfigDialog sshDialog = new SSHConfigDialog();

                // Load existing config if available
                if (_sshConfig.Count > 0)
                {
                    sshDialog.SetSSHConfig(_sshConfig);
                    Log("SSH config loaded into dialog");
                }

                // Show the dialog
                if (sshDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _sshConfig = sshDialog.GetSSHConfig();
                    Log("SSH config dialog returned OK");

                    // Show configured host info in status
                    if (_sshConfig.ContainsKey("host") && _sshConfig.ContainsKey("username"))
                    {
                        UpdateStatus($"SSH configuration updated: {_sshConfig["username"]}@{_sshConfig["host"]}");
                        Log($"SSH configuration updated: {_sshConfig["username"]}@{_sshConfig["host"]}");
                    }
                    else
                    {
                        UpdateStatus("SSH configuration updated.");
                        Log("SSH configuration updated");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnConfigureSSH_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Error configuring SSH: {ex.Message}", "SSH Configuration Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnBrowseRepo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("BtnBrowseRepo_Click started");
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select Repository Location";
                    folderDialog.ShowNewFolderButton = true;

                    System.Windows.Forms.DialogResult result = folderDialog.ShowDialog();

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        txtLocalRepoPath.Text = folderDialog.SelectedPath;
                        Log("Repository path selected: " + folderDialog.SelectedPath);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnBrowseRepo_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Error browsing for repository: {ex.Message}", "Browse Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Backup Source Management

        private void BtnAddSource_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("BtnAddSource_Click started");
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select Folder to Back Up";

                    System.Windows.Forms.DialogResult result = folderDialog.ShowDialog();

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        // Add the folder to the list if it's not already there
                        if (!listBackupSources.Items.Contains(folderDialog.SelectedPath))
                        {
                            listBackupSources.Items.Add(folderDialog.SelectedPath);
                            Log("Backup source added: " + folderDialog.SelectedPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnAddSource_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Error adding source: {ex.Message}", "Add Source Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnRemoveSource_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("BtnRemoveSource_Click started");
                if (listBackupSources.SelectedItem != null)
                {
                    listBackupSources.Items.Remove(listBackupSources.SelectedItem);
                    Log("Backup source removed: " + listBackupSources.SelectedItem);
                }
                else
                {
                    MessageBox.Show("Please select a source to remove.", "Selection Required",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                    Log("No backup source selected for removal");
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnRemoveSource_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Error removing source: {ex.Message}", "Remove Source Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Borg Operations

        private string GetRepositoryPath()
        {
            try
            {
                Log("GetRepositoryPath started");
                if (radioLocalRepository.IsChecked == true)
                {
                    Log("Local repository path: " + txtLocalRepoPath.Text);
                    return txtLocalRepoPath.Text;
                }
                else
                {
                    // For SSH repository
                    if (_sshConfig.ContainsKey("username") && _sshConfig.ContainsKey("host") && _sshConfig.ContainsKey("repo_path"))
                    {
                        string repoPath = $"{_sshConfig["username"]}@{_sshConfig["host"]}:{_sshConfig["repo_path"]}";
                        Log("SSH repository path: " + repoPath);
                        return repoPath;
                    }
                    else
                    {
                        throw new InvalidOperationException("SSH configuration is incomplete. Please configure SSH settings.");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in GetRepositoryPath: {ex.Message}\n{ex.StackTrace}");
                throw;
            }
        }

        private string GetArchiveName()
        {
            try
            {
                Log("GetArchiveName started");
                string archiveName = txtArchiveName.Text.Trim();

                if (string.IsNullOrEmpty(archiveName))
                {
                    archiveName = Environment.MachineName;
                }

                // Append timestamp to make the archive name unique
                string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                string fullArchiveName = $"{archiveName}-{timestamp}";
                Log("Archive name: " + fullArchiveName);
                return fullArchiveName;
            }
            catch (Exception ex)
            {
                Log($"Exception in GetArchiveName: {ex.Message}\n{ex.StackTrace}");
                throw;
            }
        }

        private string GetCompressionString()
        {
            try
            {
                Log("GetCompressionString started");
                string compressionType = "";
                int level = (int)sliderCompressionLevel.Value;

                // Get the selected ComboBoxItem content
                if (comboCompressionType.SelectedItem is ComboBoxItem selectedItem)
                {
                    compressionType = selectedItem.Content.ToString();
                }

                switch (compressionType)
                {
                    case "None": return "none";
                    case "Fast (Lz4)": return $"lz4,{level}";
                    case "Moderate (Zstd)": return $"zstd,{level}";
                    case "Maximum (zlib)": return $"zlib,{level}";
                    default: return "lz4,6";
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in GetCompressionString: {ex.Message}\n{ex.StackTrace}");
                throw;
            }
        }

        private string GetChunkerParams()
        {
            try
            {
                Log("GetChunkerParams started");
                string deduplication = "";

                if (comboDeduplication.SelectedItem is ComboBoxItem selectedItem)
                {
                    deduplication = selectedItem.Content.ToString();
                }

                switch (deduplication)
                {
                    case "Small chunks (better for small files)": return "10,23,16,4095";
                    case "Medium chunks (balanced)": return "12,22,18,4095";
                    case "Large chunks (better for large files)": return "19,23,21,4095";
                    default: return "10,23,16,4095";
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in GetChunkerParams: {ex.Message}\n{ex.StackTrace}");
                throw;
            }
        }

        private void BtnCreateBackup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("BtnCreateBackup_Click started");
                // Validation
                if (listBackupSources.Items.Count == 0)
                {
                    MessageBox.Show("Please add at least one source folder to back up.",
                                    "No Sources", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("No backup sources added");
                    return;
                }

                if (radioLocalRepository.IsChecked == true && string.IsNullOrWhiteSpace(txtLocalRepoPath.Text))
                {
                    MessageBox.Show("Please select a local repository path.",
                                    "Missing Repository", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("Local repository path not selected");
                    return;
                }

                if (radioSSHRepository.IsChecked == true && _sshConfig.Count == 0)
                {
                    MessageBox.Show("Please configure SSH settings for the remote repository.",
                                    "SSH Configuration", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("SSH settings not configured");
                    return;
                }

                // Get repository path
                string repoPath = GetRepositoryPath();

                // Create archive name
                string archiveName = GetArchiveName();
                string fullArchivePath = $"{repoPath}::{archiveName}";

                UpdateStatus($"Creating backup: {archiveName}");
                Log($"Creating backup: {archiveName}");

                // Build command arguments
                List<string> args = new List<string> { "create" };

                // Add compression if enabled
                if (chkCompression.IsChecked == true)
                {
                    args.Add("--compression");
                    args.Add(GetCompressionString());
                    Log("Compression added: " + GetCompressionString());
                }

                // Add chunker params for deduplication
                args.Add("--chunker-params");
                args.Add(GetChunkerParams());
                Log("Chunker params added: " + GetChunkerParams());

                // Add one-filesystem option if checked
                if (chkOneFileSystem.IsChecked == true)
                {
                    args.Add("--one-file-system");
                    Log("One-file-system option added");
                }

                // Add archive path
                args.Add(fullArchivePath);
                Log("Full archive path: " + fullArchivePath);

                // Add source paths
                foreach (string source in listBackupSources.Items)
                {
                    args.Add(source);
                    Log("Source path added: " + source);
                }

                // Execute backup command asynchronously
                Task.Run(() =>
                {
                    try
                    {
                        var result = ExecuteBorgCommand(args);
                        Log("ExecuteBorgCommand called for backup");

                        Dispatcher.Invoke(() =>
                        {
                            if (result.ContainsKey("exitCode") && (int)result["exitCode"] == 0)
                            {
                                UpdateStatus("Backup completed successfully.");
                                MessageBox.Show("Backup completed successfully.",
                                            "Backup Success", MessageBoxButton.OK, MessageBoxImage.Information);
                                Log("Backup completed successfully");
                            }
                            else
                            {
                                string errorMsg = result.ContainsKey("stderr") ? result["stderr"].ToString() : "Unknown error";
                                UpdateStatus("Backup failed: " + errorMsg, true);
                                HandleRepositoryError(errorMsg); // Use new error handler
                                Log("Backup failed: " + errorMsg);
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            UpdateStatus("Error during backup: " + ex.Message, true);
                            MessageBox.Show("Error during backup: " + ex.Message,
                                        "Backup Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            Log($"Exception during backup: {ex.Message}\n{ex.StackTrace}");
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnCreateBackup_Click: {ex.Message}\n{ex.StackTrace}");
                UpdateStatus("Error: " + ex.Message, true);
                MessageBox.Show("Error: " + ex.Message,
                            "Operation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnListArchives_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("BtnListArchives_Click started");
                // Validate repository settings
                if (radioLocalRepository.IsChecked == true && string.IsNullOrWhiteSpace(txtLocalRepoPath.Text))
                {
                    MessageBox.Show("Please select a local repository path.",
                                    "Missing Repository", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("Local repository path not selected");
                    return;
                }

                if (radioSSHRepository.IsChecked == true && _sshConfig.Count == 0)
                {
                    MessageBox.Show("Please configure SSH settings for the remote repository.",
                                    "SSH Configuration", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("SSH settings not configured");
                    return;
                }

                // Get repository path
                string repoPath = GetRepositoryPath();

                UpdateStatus($"Listing archives in repository: {repoPath}");
                Log($"Listing archives in repository: {repoPath}");

                // Execute list command asynchronously
                Task.Run(() =>
                {
                    try
                    {
                        var result = ExecuteBorgCommand(new List<string> { "list", repoPath });
                        Log("ExecuteBorgCommand called for list archives");

                        Dispatcher.Invoke(() =>
                        {
                            if (result.ContainsKey("exitCode") && (int)result["exitCode"] == 0)
                            {
                                string output = result.ContainsKey("stdout") ? result["stdout"].ToString() : "";

                                if (string.IsNullOrWhiteSpace(output))
                                {
                                    UpdateStatus("Repository exists but contains no archives.");
                                    Log("Repository exists but contains no archives");
                                }
                                else
                                {
                                    UpdateStatus("Archives in repository:");
                                    UpdateStatus(output);
                                    Log("Archives listed: " + output);
                                }
                            }
                            else
                            {
                                string errorMsg = result.ContainsKey("stderr") ? result["stderr"].ToString() : "Unknown error";
                                UpdateStatus("Failed to list archives: " + errorMsg, true);
                                HandleRepositoryError(errorMsg); // Use new error handler
                                Log("Failed to list archives: " + errorMsg);
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            UpdateStatus("Error listing archives: " + ex.Message, true);
                            MessageBox.Show("Error listing archives: " + ex.Message,
                                        "List Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            Log($"Exception in BtnListArchives_Click (inner): {ex.Message}\n{ex.StackTrace}");
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnListArchives_Click: {ex.Message}\n{ex.StackTrace}");
                UpdateStatus("Error: " + ex.Message, true);
                MessageBox.Show("Error: " + ex.Message,
                            "Operation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExtractArchive_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Log("BtnExtractArchive_Click started");
                // Validate repository settings
                if (radioLocalRepository.IsChecked == true && string.IsNullOrWhiteSpace(txtLocalRepoPath.Text))
                {
                    MessageBox.Show("Please select a local repository path.",
                                    "Missing Repository", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("Local repository path not selected");
                    return;
                }

                if (radioSSHRepository.IsChecked == true && _sshConfig.Count == 0)
                {
                    MessageBox.Show("Please configure SSH settings for the remote repository.",
                                    "SSH Configuration", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Log("SSH settings not configured");
                    return;
                }

                // First, list archives to select from
                string repoPath = GetRepositoryPath();

                UpdateStatus("Retrieving archive list...");
                Log("Retrieving archive list...");

                try
                {
                    var listResult = ExecuteBorgCommand(new List<string> { "list", "--short", repoPath });
                    Log("ExecuteBorgCommand called for list archives");

                    if (listResult.ContainsKey("exitCode") && (int)listResult["exitCode"] == 0 &&
                        listResult.ContainsKey("stdout"))
                    {
                        string[] archives = listResult["stdout"].ToString()
                            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                        if (archives.Length == 0)
                        {
                            MessageBox.Show("No archives found in the repository.",
                                            "No Archives", MessageBoxButton.OK, MessageBoxImage.Information);
                            Log("No archives found in the repository");
                            return;
                        }

                        // Show archive selection dialog
                        var selectDialog = new ArchiveSelectDialog(archives);
                        selectDialog.Owner = this;
                        bool? dialogResult = selectDialog.ShowDialog();

                        if (dialogResult == true && !string.IsNullOrEmpty(selectDialog.SelectedArchive))
                        {
                            // Ask for extraction destination
                            using (var folderDialog = new FolderBrowserDialog())
                            {
                                folderDialog.Description = "Select Extraction Destination";

                                System.Windows.Forms.DialogResult folderResult = folderDialog.ShowDialog();

                                if (folderResult == System.Windows.Forms.DialogResult.OK)
                                {
                                    string extractPath = folderDialog.SelectedPath;
                                    string fullArchivePath = $"{repoPath}::{selectDialog.SelectedArchive}";

                                    UpdateStatus($"Extracting archive {selectDialog.SelectedArchive} to {extractPath}");
                                    Log($"Extracting archive {selectDialog.SelectedArchive} to {extractPath}");

                                    // Execute extract command asynchronously
                                    Task.Run(() =>
                                    {
                                        try
                                        {
                                            var result = ExecuteBorgCommand(new List<string> {
                                                "extract",
                                                "--progress",
                                                fullArchivePath
                                            }, extractPath);
                                            Log("ExecuteBorgCommand called for extract archive");

                                            Dispatcher.Invoke(() =>
                                            {
                                                if (result.ContainsKey("exitCode") && (int)result["exitCode"] == 0)
                                                {
                                                    UpdateStatus("Extraction completed successfully.");
                                                    MessageBox.Show("Archive extracted successfully.",
                                                                "Extraction Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                                                    Log("Extraction completed successfully");
                                                }
                                                else
                                                {
                                                    string errorMsg = result.ContainsKey("stderr") ? result["stderr"].ToString() : "Unknown error";
                                                    UpdateStatus("Extraction failed: " + errorMsg, true);
                                                    HandleRepositoryError(errorMsg); // Use new error handler
                                                    Log("Extraction failed: " + errorMsg);
                                                }
                                            });
                                        }
                                        catch (Exception ex)
                                        {
                                            Dispatcher.Invoke(() =>
                                            {
                                                UpdateStatus("Error during extraction: " + ex.Message, true);
                                                MessageBox.Show("Error during extraction: " + ex.Message,
                                                            "Extraction Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                                Log($"Exception during extraction: {ex.Message}\n{ex.StackTrace}");
                                            });
                                        }
                                    });
                                }
                            }
                        }
                    }
                    else
                    {
                        string errorMsg = listResult.ContainsKey("stderr") ? listResult["stderr"].ToString() : "Unknown error";
                        UpdateStatus("Failed to list archives: " + errorMsg, true);
                        HandleRepositoryError(errorMsg); // Use new error handler
                        Log("Failed to list archives: " + errorMsg);
                    }
                }
                catch (Exception ex)
                {
                    Log($"Exception in BtnExtractArchive_Click (inner): {ex.Message}\n{ex.StackTrace}");
                    UpdateStatus("Error listing archives: " + ex.Message, true);
                    MessageBox.Show("Error listing archives: " + ex.Message, "List Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnExtractArchive_Click: {ex.Message}\n{ex.StackTrace}");
                UpdateStatus("Error: " + ex.Message, true);
                MessageBox.Show("Error: " + ex.Message,
                            "Operation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Utility Methods

        private Dictionary<string, object> ExecuteBorgCommand(List<string> args, string workingDirectory = null)
        {
            var result = new Dictionary<string, object>();

            try
            {
                Log($"ExecuteBorgCommand started with args: {string.Join(" ", args)}");

                // Capture UI state on the UI thread before proceeding
                bool isSSH = false;
                Dictionary<string, string> sshConfigCopy = null;

                if (Dispatcher.CheckAccess())
                {
                    // We're on the UI thread, safe to access UI properties directly
                    isSSH = radioSSHRepository != null && radioSSHRepository.IsChecked == true;
                    if (isSSH && _sshConfig.Count > 0)
                    {
                        sshConfigCopy = new Dictionary<string, string>(_sshConfig);
                    }
                }
                else
                {
                    // We're on a background thread, use Dispatcher to safely access UI
                    Dispatcher.Invoke(() =>
                    {
                        isSSH = radioSSHRepository != null && radioSSHRepository.IsChecked == true;
                        if (isSSH && _sshConfig.Count > 0)
                        {
                            sshConfigCopy = new Dictionary<string, string>(_sshConfig);
                        }
                    });
                }

                var processInfo = new ProcessStartInfo
                {
                    FileName = _borgPath,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                // Set working directory if specified
                if (!string.IsNullOrEmpty(workingDirectory))
                {
                    processInfo.WorkingDirectory = workingDirectory;
                    Log($"Working directory set to: {workingDirectory}");
                }

                // Configure SSH environment variables if using SSH repository
                if (isSSH && sshConfigCopy != null && sshConfigCopy.Count > 0)
                {
                    if (sshConfigCopy.ContainsKey("port"))
                    {
                        processInfo.EnvironmentVariables["BORG_RSH"] = $"ssh -p {sshConfigCopy["port"]}";

                        if (sshConfigCopy.ContainsKey("auth_method") && sshConfigCopy["auth_method"] == "key" && sshConfigCopy.ContainsKey("key_path"))
                        {
                            processInfo.EnvironmentVariables["BORG_RSH"] += $" -i \"{sshConfigCopy["key_path"]}\"";
                        }

                        Log($"BORG_RSH set to: {processInfo.EnvironmentVariables["BORG_RSH"]}");
                    }

                    // Add password if using password auth and remember password is enabled
                    if (sshConfigCopy.ContainsKey("auth_method") && sshConfigCopy["auth_method"] == "password" && sshConfigCopy.ContainsKey("password"))
                    {
                        processInfo.EnvironmentVariables["BORG_PASSPHRASE"] = sshConfigCopy["password"];
                        Log("BORG_PASSPHRASE set");
                    }
                }

                // Add arguments - using Arguments for better compatibility
                processInfo.Arguments = string.Join(" ", args.Select(arg => arg.Contains(" ") ? $"\"{arg}\"" : arg));
                Log($"Process arguments: {processInfo.Arguments}");

                // Execute process
                using (Process process = new Process())
                {
                    process.StartInfo = processInfo;

                    // Capture output
                    StringBuilder outputBuilder = new StringBuilder();
                    StringBuilder errorBuilder = new StringBuilder();

                    process.OutputDataReceived += (sender, e) =>
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(e.Data))
                            {
                                outputBuilder.AppendLine(e.Data);
                                Log($"Process stdout: {e.Data}");

                                // Update UI with progress information - avoid dispatcher for initial version check
                                if (args.Count == 1 && args[0] == "--version")
                                {
                                    // Skip UI updates for version check
                                }
                                else if (e.Data.Contains("%"))
                                {
                                    try
                                    {
                                        Dispatcher?.Invoke(() =>
                                        {
                                            try
                                            {
                                                UpdateStatus(e.Data, false, true); // Progress info
                                            }
                                            catch (Exception ex)
                                            {
                                                Log($"Inner UpdateStatus exception: {ex.Message}");
                                            }
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Dispatcher invoke exception: {ex.Message}");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"OutputDataReceived exception: {ex.Message}");
                        }
                    };

                    process.ErrorDataReceived += (sender, e) =>
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(e.Data))
                            {
                                errorBuilder.AppendLine(e.Data);
                                Log($"Process stderr: {e.Data}");

                                // Skip UI updates for version check to avoid potential threading issues
                                if (args.Count == 1 && args[0] == "--version")
                                {
                                    // Skip UI updates for version check
                                }
                                else if (!e.Data.Contains("passphrase") && !e.Data.Contains("Keeping"))
                                {
                                    try
                                    {
                                        Dispatcher?.Invoke(() =>
                                        {
                                            try
                                            {
                                                UpdateStatus(e.Data, false, true);
                                            }
                                            catch (Exception ex)
                                            {
                                                Log($"Inner UpdateStatus exception: {ex.Message}");
                                            }
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Dispatcher invoke exception: {ex.Message}");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"ErrorDataReceived exception: {ex.Message}");
                        }
                    };

                    Log("Starting process");

                    try
                    {
                        bool processStarted = process.Start();
                        if (!processStarted)
                        {
                            Log("Failed to start process");
                            result["exitCode"] = -1;
                            result["stderr"] = "Failed to start process";
                            return result;
                        }

                        Log("Process started successfully");

                        // Begin asynchronous reading
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();

                        Log("Async reading started");

                        // Wait for the process to exit with timeout
                        bool exited = process.WaitForExit(30000); // 30-second timeout
                        if (!exited)
                        {
                            Log("Process did not exit within timeout period. Attempting to kill.");
                            try
                            {
                                process.Kill();
                                Log("Process killed after timeout");
                            }
                            catch (Exception ex)
                            {
                                Log($"Failed to kill process: {ex.Message}");
                            }

                            result["exitCode"] = -1;
                            result["stderr"] = "Process timed out";
                            return result;
                        }

                        Log($"Process exited with code: {process.ExitCode}");

                        // Store results
                        result["exitCode"] = process.ExitCode;
                        result["stdout"] = outputBuilder.ToString().TrimEnd('\r', '\n');
                        result["stderr"] = errorBuilder.ToString().TrimEnd('\r', '\n');

                        Log("Results stored and returning from ExecuteBorgCommand");
                    }
                    catch (Exception ex)
                    {
                        Log($"Exception starting or running process: {ex.Message}\n{ex.StackTrace}");
                        result["exitCode"] = -1;
                        result["stderr"] = $"Error executing process: {ex.Message}";
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                Log($"Exception in ExecuteBorgCommand: {ex.Message}\n{ex.StackTrace}");
                result["exitCode"] = -1;
                result["stderr"] = $"ExecuteBorgCommand error: {ex.Message}";
                return result;
            }
            finally
            {
                Log("Exiting ExecuteBorgCommand");
            }
        }

        private void HandleRepositoryError(string errorMsg)
        {
            // Check for common permission-related error patterns
            if (errorMsg.Contains("Failed to create/acquire the lock") ||
                errorMsg.Contains("Permission denied") ||
                errorMsg.Contains("No usable temporary directory name found") ||
                errorMsg.Contains("Access is denied"))
            {
                string userFriendlyMessage = "You don't have permission to access this repository location.\n\n" +
                                            "Please either:\n" +
                                            "• Choose a different folder in your user directory\n" +
                                            "• Run the application as administrator\n\n" +
                                            "Technical details: " + errorMsg;

                MessageBox.Show(userFriendlyMessage, "Permission Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                UpdateStatus("Permission error accessing repository. Choose a different location or run as administrator.", true);
                Log("Repository permission error detected");
            }
            else if (errorMsg.Contains("repository does not exist") ||
                     errorMsg.Contains("not a valid repository") ||
                     errorMsg.Contains("does not exist"))
            {
                string userFriendlyMessage = "The repository doesn't exist or hasn't been initialized.\n\n" +
                                             "You may need to initialize the repository first with the 'borg init' command.";
                MessageBox.Show(userFriendlyMessage, "Repository Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                UpdateStatus("Repository needs to be initialized first.", true);
                Log("Repository initialization required");
            }
            else
            {
                // For other errors, display the original message
                MessageBox.Show(errorMsg, "Repository Error", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus("Error: " + errorMsg, true);
            }
        }

        private void UpdateStatus(string message, bool isError = false, bool isProgress = false)
        {
            try
            {
                if (txtStatus.Dispatcher.CheckAccess())
                {
                    string timestamp = DateTime.Now.ToString("HH:mm:ss");
                    string line = isProgress ? message : $"[{timestamp}] {message}";

                    // Append the message
                    if (isProgress && txtStatus.Text.Length > 0)
                    {
                        // Replace last line if it's progress info
                        string[] lines = txtStatus.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                        if (lines.Length > 0)
                        {
                            lines[lines.Length - 1] = line;
                            txtStatus.Text = string.Join(Environment.NewLine, lines);
                        }
                        else
                        {
                            txtStatus.AppendText(line + Environment.NewLine);
                        }
                    }
                    else
                    {
                        txtStatus.AppendText(line + Environment.NewLine);
                    }

                    txtStatus.ScrollToEnd();
                }
                else
                {
                    txtStatus.Dispatcher.Invoke(() => UpdateStatus(message, isError, isProgress));
                }
            }
            catch (Exception ex)
            {
                // Log status update errors
                Log($"Error in UpdateStatus: {ex.Message}");
                // Ignore errors in status updates
            }
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                txtStatus.Clear();
                UpdateStatus("Log cleared.");
                Log("Log cleared by user");
            }
            catch (Exception ex)
            {
                Log($"Exception in BtnClearLog_Click: {ex.Message}\n{ex.StackTrace}");
                // Continue without crashing
            }
        }

        private void Log(string message)
        {
            try
            {
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WinBorg_debug.log");
                File.AppendAllText(path, $"[{DateTime.Now:HH:mm:ss.fff}] {message}\n");
                Debug.WriteLine($"APP DEBUG: {message}");
            }
            catch (Exception ex)
            {
                // Try console output if file logging fails
                Debug.WriteLine($"APP DEBUG: Logging error: {ex.Message}");
                Debug.WriteLine($"APP DEBUG: Original message: {message}");
            }
        }

        #endregion
    }
}
