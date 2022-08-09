using AutoUpdaterDotNET;
using ExcelReporter.Properties;
using Octokit;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static System.Environment;

namespace ExcelReporter
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
#if !DEBUG
            // since debugging source always be the latest version, therefore no need to check for update in debug mode
            Program.TryUpdateApplicationAsync();
#endif

            // check if the application is need to be upgraded and copy the settings from older version to this version
            if (Settings.Default.UpgradeRequired)
            {
                Settings.Default.Upgrade();
                Settings.Default.UpgradeRequired = false;
                Settings.Default.Save();
            }

            // make some settings to be default on first launch
            Program.buildSettingsFirstLaunch();

            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new FrmMain());
        }

        /// <summary>
        /// Set some settings on first launch
        /// </summary>
        private static void buildSettingsFirstLaunch()
        {
            bool settingsChanged = false;

            if (string.IsNullOrEmpty(Settings.Default.ReportsSaveLocation))
            {
                Settings.Default.ReportsSaveLocation = Path.Combine(GetFolderPath(SpecialFolder.MyDocuments), "Reports");
                settingsChanged = true;
            }

            if (settingsChanged)
                Settings.Default.Save();
        }

        /// <summary>
        /// Try getting update information from github releases
        /// </summary>
        private static async Task TryUpdateApplicationAsync()
        {
            var client = new GitHubClient(new ProductHeaderValue(BuildConstants.GITHUB_REPOSITORY));
            var releases = await client.Repository.Release.GetAll(BuildConstants.GITHUB_USERNAME, BuildConstants.GITHUB_REPOSITORY);

            // stop the process if cannot get release information from github
            if (releases == null)
                return;

            if (releases.Count > 0)
            {
                var latestRelease = releases[0];

                // get remote version and local version
                Version remoteVersion = Helpers.GithubTagToVersion(latestRelease.TagName);
                Version localVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

                // compare two versions
                if (remoteVersion <= localVersion)
                    return;

                string downloadFileName = string.Format(BuildConstants.GITHUB_ASSET_NAME, latestRelease.TagName);

                // if there is only source files in release, then stop the updater
                if (latestRelease.Assets.Count == 0) return;

                var downloadAsset = latestRelease.Assets.First(a => a.Name.StartsWith(downloadFileName) && a.Name.EndsWith("release.zip"));

                // if both latestVersion and downloadAsset are not null
                if (remoteVersion != null)
                {
                    // create xml file for updating
                    string appCastFilePath = Path.Combine(Path.GetTempPath(), "update.xml");
                    Helpers.MakeAppCastFile(remoteVersion, downloadAsset.BrowserDownloadUrl, appCastFilePath);

                    AutoUpdater.Start(appCastFilePath);
                }
            }
        }
    }
}