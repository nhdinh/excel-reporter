﻿using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using static System.Environment;

namespace ExcelReporter
{
    internal static class Helpers
    {
        public static string SettingsSavePath = GetFolderPath(SpecialFolder.LocalApplicationData);
        public static string SettingsFolderName = "Reporter";
        public static string RecentFileSaveName = "recents.json";
        public static string AppConfigSaveName = "config.json";
        public static string ReportOptionSaveNamePrfx = "reop_";

        public static void ConfigNLog()
        {
            var config = new NLog.Config.LoggingConfiguration();

            var logFile = new NLog.Targets.FileTarget("logfile") { FileName = "app.log" };
            var logConsole = new NLog.Targets.ConsoleTarget("logconsole");

#if (DEBUG)
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logConsole);
#elif (RELEASE)
			config.AddRule(LogLevel.Info, LogLevel.Fatal, logConsole);
#endif
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logFile);

            NLog.LogManager.Configuration = config;
        }

        internal static string[] FetchReportOptions()
        {
            string savePath = Path.Combine(SettingsSavePath, SettingsFolderName);
            if (!Directory.Exists(savePath))
                return new string[] { };
            else
            {
                IEnumerable<string> files = Directory.EnumerateFiles(savePath, ReportOptionSaveNamePrfx + "*.json");
                List<string> reportOptionIdList = new List<string>();
                foreach (var filePath in files)
                {
                    var reportOptionId = Path.GetFileName(filePath).Replace(ReportOptionSaveNamePrfx, "").Replace(".json", "");

                    reportOptionIdList.Add(reportOptionId);
                }

                return reportOptionIdList.ToArray();
            }
        }

        internal static void SaveJson(string fileName, object o)
        {
            string savePath = Path.Combine(SettingsSavePath, SettingsFolderName);
            if (!Directory.Exists(savePath))
                Directory.CreateDirectory(savePath);

            JsonSerializer serializer = new JsonSerializer();

            using (var fs = new StreamWriter(Path.Combine(savePath, fileName)))
            using (JsonWriter writer = new JsonTextWriter(fs))
                serializer.Serialize(writer, o);
        }

        internal static void SaveRecentFile(IList<WorkFileInfo> workingFilesData)
        {
            SaveJson(RecentFileSaveName, workingFilesData);
        }

        internal static IList<WorkFileInfo> LoadRecentFile()
        {
            // load from recent files data
            var recentsFilePath = Path.Combine(SettingsSavePath, SettingsFolderName, RecentFileSaveName);

            if (File.Exists(recentsFilePath))
            {
                var _workingFiles = JsonConvert.DeserializeObject<List<WorkFileInfo>>(File.ReadAllText(recentsFilePath));

                return _workingFiles;
            }

            return new List<WorkFileInfo>();
        }

        internal static void SaveAppConfig(AppConfig config)
        {
            SaveJson(AppConfigSaveName, config);
        }

        internal static AppConfig LoadAppConfigFromJson()
        {
            var appConfigFilePath = Path.Combine(SettingsSavePath, SettingsFolderName, AppConfigSaveName);
            if (File.Exists(appConfigFilePath))
            {
                var _appConfig = JsonConvert.DeserializeObject<AppConfig>(File.ReadAllText(appConfigFilePath));

                return _appConfig;
            }

            return new AppConfig();
        }

        internal static void SaveReportOption(ReportOption option)
        {
            SaveJson(ReportOptionSaveNamePrfx + option.Id + ".json", option);
        }

        internal static ReportOption LoadReportOption(string reportOptionId)
        {
            var optionFile = Path.Combine(SettingsSavePath, SettingsFolderName);
            optionFile = Path.Combine(optionFile, ReportOptionSaveNamePrfx + reportOptionId + ".json");

            if (File.Exists(optionFile))
            {
                var reportOption = JsonConvert.DeserializeObject<ReportOption>(File.ReadAllText(optionFile));

                return reportOption;
            }

            return null;
        }

        internal static string MakeNewReportOptionId(string text)
        {
            string illegal = "\"M\"\\a/ry/ h**ad:>> a\\/:*?\"| li*tt|le|| la\"mb.?";

            foreach (char c in text)
            {
                illegal = illegal.Replace(c.ToString(), "");
            }

            return text.Trim();
        }

        internal static void DeleteReportOption(string optionId)
        {
            var optionFile = Path.Combine(SettingsSavePath, SettingsFolderName);
            optionFile = Path.Combine(optionFile, ReportOptionSaveNamePrfx + optionId + ".json");

            if (File.Exists(optionFile))
                try
                {
                    File.Delete(optionFile);
                }
                catch { throw; }
        }

        internal static void DeleteReportOption(ReportOption option)
        {
            DeleteReportOption(option.Id);
        }

        /// <summary>
        /// Make appCast file for AutoUpdater.NET and save it to somewhere specified by appCastFilePath.
        /// The format of appCast is XML, and contains needed information for AutoUpdater.NET to update the application.
        /// </summary>
        /// <param name="latestVersion">the latest remote version</param>
        /// <param name="downloadUrl">asset download URL</param>
        /// <param name="appCastFilePath">appCast file location</param>
        public static void MakeAppCastFile(Version latestVersion, string downloadUrl, string appCastFilePath)
        {
            using (XmlWriter writer = XmlWriter.Create(appCastFilePath))
            {
                writer.WriteStartElement("item");
                writer.WriteElementString("version", latestVersion.ToString());
                writer.WriteElementString("url", downloadUrl);
                writer.WriteElementString("mandator", "true");
            }
        }

        /// <summary>
        /// Convert github tag in string format to Version instance.
        /// Github tag has format of v0.0.1. This is the release tag of the repository.
        /// This function will parse the tag and create an instance of Version with those version number.
        ///
        /// If input tag is null or empty or not in specified string, return null
        /// </summary>
        /// <param name="tagName">a github tag string in format v*.*.*</param>
        /// <returns>instance of Version</returns>
        public static Version GithubTagToVersion(string tagName)
        {
            if (string.IsNullOrWhiteSpace(tagName)) return null;

            string strRegex = @"^v([0-9]+)\.([0-9]+)\.([0-9]+)((\.[0-9]+)?)$";
            var regexMatch = new Regex(strRegex);
            if (regexMatch.IsMatch(tagName))
            {
                tagName = tagName.Replace("v", "");
                int[] versions = tagName.Split('.').Select(v => int.Parse(v)).ToArray();

                if (versions.Length == 2)
                    return new Version(versions[0], versions[1], 0);
                else if (versions.Length == 3)
                    return new Version(versions[0], versions[1], versions[2]);
                else if (versions.Length == 4)
                    return new Version(versions[0], versions[1], versions[2], versions[3]);
                else return null;
            }

            return null;
        }

        /// <summary>
        /// Get RuntimeIdentifier of the current process
        /// </summary>
        /// <returns>a string represent for RuntimeIdentifier</returns>
        public static string GetRuntimeIdentifier()
        {
            string os = "";

            // check if the current OS is windows?
            if (Environment.OSVersion.Platform == PlatformID.Win32NT) os = "win";

            if (RuntimeInformation.OSArchitecture == Architecture.X64)
                return os + "-x64";
            else if (RuntimeInformation.OSArchitecture == Architecture.X86)
                return os + "-x86";

            return "unknown";
        }
    }
}