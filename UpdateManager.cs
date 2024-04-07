using System;
using System.Diagnostics;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace LSS
{
    public class UpdateManager
    {
        public async Task<bool> CheckForUpdates()
        {
            try
            {
                // Get the latest version information from the server
                Version latestVersion = await GetLatestVersion();
                // Get the current version of the application
                Version currentVersion = Assembly.GetEntryAssembly().GetName().Version;
                String zipFilename = "";

                // Log the current and latest versions
                Console.WriteLine($"Current version: {currentVersion}");
                Console.WriteLine($"Latest version: {latestVersion}");

                if (latestVersion > currentVersion)
                {
                    // Get the release notes from the server
                    string releaseNotes = await DownloadReleaseNotes("https://ajhsoftware.co.uk/lss_app/release_notes.txt");

                    // Get new zip filename
                    zipFilename = "LSS-v" + latestVersion.ToString().Replace('.', '_') + ".zip";

                    // Display the update dialog with release notes and clickable link
                    await ShowUpdateDialog(releaseNotes, currentVersion, latestVersion, zipFilename);

                    // We assume the user clicked "Update" and initiated the update process
                    // You can add more logic here to handle the actual update process
                    return true;
                }
                else
                {
                    MessageBox.Show("No updates available.\nCurrent version: " + currentVersion + "\nLatest version: " + latestVersion, "Update Check", MessageBoxButton.OK, MessageBoxImage.Information);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while checking for updates: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private async Task<Version> GetLatestVersion()
        {
            using (var httpClient = new HttpClient())
            {
                string versionString = await httpClient.GetStringAsync("https://ajhsoftware.co.uk/lss_app/version.txt");
                return new Version(versionString.Trim());
            }
        }



        private async Task<string> DownloadReleaseNotes(string releaseNotesUrl)
        {
            using (var httpClient = new HttpClient())
            {
                return await httpClient.GetStringAsync(releaseNotesUrl);
            }
        }

        private async Task ShowUpdateDialog(string releaseNotes, Version currentVersion, Version latestVersion, string zipFilename)
        {
            // Display the release notes to the user in a custom dialog
            var dialog = new Window
            {
                Title = "Update Available",
                Width = 650,
                Height = 400,
                Content = new StackPanel
                {
                    Children =
            {
                new TextBlock
                {
                    Text = $"Current Version: {currentVersion}",
                    FontSize = 14,
                    Margin = new Thickness(5),
                },
                new TextBlock
                {
                    Text = $"Latest Version: {latestVersion}",
                    FontSize = 14,
                    Margin = new Thickness(5),
                },
                new TextBlock
                {
                    Text = "Release Notes:",
                    FontSize = 18,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(5),
                },
                new RichTextBox
                {
                    IsReadOnly = true,
                    Width = 580,
                    Height = 200,
                    Margin = new Thickness(5),
                    Document = new FlowDocument(new Paragraph(new Run(releaseNotes)))
                }
            }
                }
            };

            var linkTextBlock = new TextBlock
            {
                Text = "Download Update",
                Margin = new Thickness(5),
                TextWrapping = TextWrapping.Wrap,
                Foreground = System.Windows.Media.Brushes.Blue,
                Cursor = System.Windows.Input.Cursors.Hand
            };

            linkTextBlock.MouseLeftButtonDown += (sender, e) =>
            {
                Process.Start(new ProcessStartInfo("https://ajhsoftware.co.uk/lss_app/" + zipFilename) { UseShellExecute = true });
            };

            ((StackPanel)dialog.Content).Children.Add(linkTextBlock);

            await Application.Current.Dispatcher.InvokeAsync(() => dialog.ShowDialog());
        }
    }

}