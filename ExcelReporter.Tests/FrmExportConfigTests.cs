using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.IO;

namespace ExcelReporter.Tests
{
    [TestClass]
    public class FrmExportConfigTests
    {
        [TestMethod]
        public void MakeDailyReportsTestReturnNullIfReportSourceIsNull()
        {
            var reportKeys = new ReportKey[] {
            new ReportKey(new DateTime(2018,1,1), "PN001", 0.1, "Inspector 1", "Shift 1"),
            new ReportKey(new DateTime(2018,1,2), "PN001", 0.1, "Inspector 1", "Shift 1")
            };

            var sourceReport = new DataTable();

            var frm = new FrmExportConfig();
            frm.frmExportConfig_Load(frm, null);
            object result = frm.MakeDailyReports(reportKeys, sourceReport);

            Assert.IsTrue(result == null);
        }

        [TestMethod]
        public void CheckReportSavingLocationTest()
        {
            const string reportFolderName = "fname";
            var reportSaveLocation = Path.GetTempPath();

            var frm = new FrmExportConfig();
            frm.checkReportSavingLocation(reportFolderName, reportSaveLocation);

            var expectedFolder = Path.Combine(reportSaveLocation, reportFolderName);
            Assert.IsTrue(Directory.Exists(expectedFolder));

            Directory.Delete(expectedFolder);
            Assert.IsFalse(Directory.Exists(expectedFolder));
        }
    }
}