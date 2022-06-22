using ExcelReporter.Exceptions;
using ExcelReporter.Properties;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using static System.Environment;

namespace ExcelReporter
{
    internal static class Helpers
    {
        public static string SettingsSavePath = GetFolderPath(SpecialFolder.LocalApplicationData);
        public static string SettingsFolderName = "Reporter";
        public static string RecentFileSaveName = $"recents.json";
        public static string AppConfigSaveName = "config.json";
        public static string ReportOptionSaveNamePrefix = "reop_";

        public static void ConfigNLog()
        {
            var config = new NLog.Config.LoggingConfiguration();

            var logFile = new NLog.Targets.FileTarget("logfile") { FileName = "app.log" };
            var logConsole = new NLog.Targets.ConsoleTarget($"logconsole");

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
            var savePath = Path.Combine(SettingsSavePath, SettingsFolderName);
            if (!Directory.Exists(savePath))
                return new string[] { };
            else
            {
                var files = Directory.EnumerateFiles(savePath, ReportOptionSaveNamePrefix + "*.json");
                var reportOptionIdList = new List<string>();
                foreach (var filePath in files)
                {
                    var reportOptionId = Path.GetFileName(filePath).Replace(ReportOptionSaveNamePrefix, "").Replace(".json", "");

                    reportOptionIdList.Add(reportOptionId);
                }

                return reportOptionIdList.ToArray();
            }
        }

        internal static void SaveJson(string fileName, object o)
        {
            var savePath = Path.Combine(SettingsSavePath, SettingsFolderName);
            if (!Directory.Exists(savePath))
                Directory.CreateDirectory(savePath);

            var serializer = new JsonSerializer();

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
            var recentFilePath = Path.Combine(SettingsSavePath, SettingsFolderName, RecentFileSaveName);

            if (!File.Exists(recentFilePath)) return new List<WorkFileInfo>();
            var _workingFiles = JsonConvert.DeserializeObject<List<WorkFileInfo>>(File.ReadAllText(recentFilePath));

            return _workingFiles;

        }

        internal static void SaveReportOption(ReportOption option)
        {
            SaveJson(ReportOptionSaveNamePrefix + option.Id + ".json", option);
        }

        internal static ReportOption LoadReportOption(string reportOptionId)
        {
            var optionFile = Path.Combine(SettingsSavePath, SettingsFolderName);
            optionFile = Path.Combine(optionFile, ReportOptionSaveNamePrefix + reportOptionId + ".json");

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
            optionFile = Path.Combine(optionFile, ReportOptionSaveNamePrefix + optionId + ".json");

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

        internal static string GetReportsSaveLocation()
        {
            return !string.IsNullOrEmpty(Settings.Default.ReportsSaveLocation) ? Settings.Default.ReportsSaveLocation : Path.Combine(GetFolderPath(SpecialFolder.MyDocuments), "Reports");
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

        internal static string GetReportTemplate(string reportForm)
        {
            var template = reportForm == "201" ? @".\templates\mt201_template.xlsx" : @".\templates\mt202_template.xlsx";
            if (File.Exists(template))
                return template;
            else throw new TemplateNotFoundException("Template for report form " + reportForm + " not found");
        }

        internal static string makeReportSheetName(DateTime inspectedDate, double concentration, string partNo, string inspector, string shiftName, int page)
        {
            string sheetName;

            if (page > 1)
                sheetName = String.Format("{0:dd}_{1}_{2}_{3}{4}_P{5}", inspectedDate, concentration, partNo, inspector, shiftName, page);
            else
                sheetName = String.Format("{0:dd}_{1}_{2}_{3}{4}", inspectedDate, concentration, partNo, inspector, shiftName);

            return sheetName;
        }

        internal static string makeReportSheetName(ReportKey reportKey, int curPage)
        {
            DateTime inspDate = reportKey.InspectedDate;
            double concentration = reportKey.Concentration * 100;
            string partNo = reportKey.PartNo;
            string inspector = reportKey.Inspector.Length > 3 ? reportKey.Inspector.Remove(3) : reportKey.Inspector;
            string shift = reportKey.Shift.Replace("S", "");

            return makeReportSheetName(inspDate, concentration, partNo, inspector, shift, curPage);
        }
    }
}