using ExcelReporter.Exceptions;
using ExcelReporter.Properties;
using NLog;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

[assembly: InternalsVisibleTo("ExcelReporter.Tests")]

namespace ExcelReporter
{
    public partial class FrmExportConfig : Form
    {
        private const string NewReportOptionLabel = "New...";

        private readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly CoreApp app;
        private bool dirty = false;
        private ReportOption currentReportOption;
        private List<bool> validated = new List<bool>();
        private readonly BindingList<Part> dgBindingList = new BindingList<Part>();

        public FrmExportConfig()
        {
            InitializeComponent();

            // init app
            this.app = CoreApp.GetInstance();
            this.ReportOptionChanged += frmExportConfig_ReportOptionChanged;
        }

        public event EventHandler ReportOptionChanged;

        protected ReportOption CurrentReportOption
        {
            get => this.currentReportOption;
            private set
            {
                this.currentReportOption = value;
                this.OnReportOptionChanged(this.currentReportOption);
            }
        }

        protected virtual void OnReportOptionChanged(ReportOption option)
        {
            ReportOptionChanged?.Invoke(option, EventArgs.Empty);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (dirty)
            {
                DialogResult res = MessageBox.Show(this, @"Data is not saved. Close anyway?", @"Close?",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (res == DialogResult.Yes)
                {
                    this.DialogResult = DialogResult.Cancel;
                    this.Close();
                }
            }
            else
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            string reportOptionId = (string)this.cbSelectReportOptionId.SelectedItem;
            if (reportOptionId != NewReportOptionLabel)
                try
                {
                    Helpers.DeleteReportOption(reportOptionId);
                    this.LoadReportOptionNamesToComboBox();
                }
                catch
                {
                    MessageBox.Show(this, @"Delete failed", @"Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
        }

        private void btnSaveAndClose_Click(object sender, EventArgs e)
        {
            this.validated = new List<bool>();

            foreach (var textBox in groupBox1.Controls)
            {
                if (textBox.GetType() == typeof(TextBox))
                {
                    this.ValidateTextFields((TextBox)textBox);
                }
            }

            // return ìf any field is invalid
            if (!validated.All(x => x)) return;

            var reportOptionId = (string)this.cbSelectReportOptionId.SelectedItem;
            reportOptionId = reportOptionId?.Trim();

            if (reportOptionId == NewReportOptionLabel || string.IsNullOrEmpty(reportOptionId))
            {
                this.CurrentReportOption.Id = Helpers.MakeNewReportOptionId(this.txtCustomerName.Text);
            }

            this.CurrentReportOption.CustomerName = this.txtCustomerName.Text.Trim();
            this.CurrentReportOption.Procedure = this.txtProcedure.Text.Trim();
            this.CurrentReportOption.AcceptanceCriteria = this.txtAcceptanceCriteria.Text.Trim();

            this.CurrentReportOption.Parts = new List<Part>();
            foreach (DataGridViewRow row in this.dgParts.Rows)
            {
                if (row.Cells["PartNo"].Value == null) continue;

                var part = new Part()
                {
                    PartNo = row.Cells["PartNo"].Value != null ? (string)row.Cells["PartNo"].Value : "",
                    Description = row.Cells["Description"].Value != null ? (string)row.Cells["Description"].Value : "",
                    Coil = int.Parse(row.Cells["Coil"].Value.ToString()),
                    Yoke = int.Parse(row.Cells["Yoke"].Value.ToString()),
                    ExtendOfTest = row.Cells["ExtendOfTest"].Value != null
                        ? (string)row.Cells["ExtendOfTest"].Value
                        : "",
                };

                this.CurrentReportOption.Parts.Add(part);
            }

            // save report option
            Helpers.SaveReportOption(this.CurrentReportOption);

            // save config
            Settings.Default.LastReportOptionId = this.CurrentReportOption.Id;
            Settings.Default.Save();

            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnSaveAndExport_Click(object sender, EventArgs e)
        {
            var startDate = this.dtpStart.Value.Date;
            var endDate = this.dtpEnd.Value.Date;
            var reportForm = this.mt201.Checked ? "201" : "202";

            if (!this.validateRequestedDates(startDate, endDate))
                return;

            // make dataset from excel report
            DataSet dataset = null;
            try
            {
                dataset = this.app.MakeReportData(startDate, endDate);
            }
            catch (Exception exc)
            {
                this.logger.Error(exc);
                MessageBox.Show(this, @"Error while making report data. " + exc.Message, "Error...",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }

            if (dataset != null && dataset.Tables["source"].Rows.Count > 0)
            {
                string reportsSaveLocation = Helpers.GetReportsSaveLocation();
                string reportsFolderName = $"{startDate:yyMMdd}_{endDate:yyMMdd}";

                // check the target save folder if existed and contains files, then need asking to backup, else create new
                if (!this.checkReportSavingLocation(reportsFolderName, reportsSaveLocation))
                    return;

                var reportDateRange = dataset.Tables["report_key"].AsEnumerable()
                    .GroupBy(g => new { InspectedDate = g.Field<DateTime>("Inspected Date") })
                    .Where(g => g.Any())
                    .OrderBy(g => g.Key.InspectedDate)
                    .Select(g => new DateTime(
                        g.Key.InspectedDate.Year,
                        g.Key.InspectedDate.Month,
                        g.Key.InspectedDate.Day
                    )).ToArray();

                // loop through each date in the reportDateRange and create workbook for each date
                foreach (var reportDate in reportDateRange)
                {
                    ReportKey[] reportKeys = dataset.Tables["report_key"].AsEnumerable()
                        .GroupBy(row => new
                        {
                            InspectedDate = row.Field<DateTime>("Inspected Date"),
                            PartNo = row.Field<string>("Part No."),
                            Concentration = row.Field<double>("Concentration"),
                            Inspector = row.Field<string>("Inspector"),
                            Shift = row.Field<string>("Shift")
                        }).Where(g => g.Key.InspectedDate == reportDate)
                        .Select(o => new ReportKey(
                            o.Key.InspectedDate,
                            o.Key.PartNo,
                            o.Key.Concentration,
                            o.Key.Inspector,
                            o.Key.Shift
                        )).ToArray();

                    // make excel workbook base on input data
                    var reportWorkbook = this.MakeDailyReports(reportKeys, dataset.Tables["source"], reportForm);
                    if (reportWorkbook == null)
                    {
                        MessageBox.Show(@"No data to export", @"Error...", MessageBoxButtons.OK);
                        this.DialogResult = DialogResult.OK;
                        this.Close();
                    }

                    // else, generate the report file
                    SaveReportFile(reportWorkbook, Path.Combine(reportsSaveLocation, reportsFolderName), reportDate,
                        reportForm);
                }

                ;

                MessageBox.Show(this, $@"Exported successfully to {reportsSaveLocation}.", @"Done",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        /// <summary>
        /// Check for the saving folder of exported reports is ready or not. Return True if all OK, else False
        /// </summary>
        /// <param name="reportsSaveLocation">default saving location</param>
        /// <param name="reportsFolderName">The folder name of which contains the output report files</param>
        internal bool checkReportSavingLocation(string reportsFolderName, string reportsSaveLocation = "")
        {
            if (string.IsNullOrEmpty(reportsSaveLocation))
                reportsSaveLocation = Helpers.GetReportsSaveLocation();

            if (Directory.Exists(Path.Combine(reportsSaveLocation, reportsFolderName)))
            {
                if (Directory.GetFiles(Path.Combine(reportsSaveLocation, reportsFolderName), "*.*").Length != 0)
                {
                    DialogResult res = MessageBox.Show(this,
                        $@"Report for {reportsFolderName} already exists. Make backup and continue?",
                        @"Reports exists", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button2);
                    if (res == DialogResult.Yes)
                    {
                        try
                        {
                            // move to another folder
                            Directory.Move(Path.Combine(reportsSaveLocation, reportsFolderName),
                                Path.Combine(reportsSaveLocation,
                                    $"{reportsFolderName}_backup{DateTime.Now:yyMMddHHmmss}"));

                            // recreate the report folder
                            Directory.CreateDirectory(Path.Combine(reportsSaveLocation, reportsFolderName));
                        }
                        catch
                        {
                            MessageBox.Show(
                                @"There is an error while creating the backup, close any opening related files and try again.",
                                @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                    }
                    else
                    {
                        return false; // cancel
                    }
                }
            }
            else
                // create directory
                Directory.CreateDirectory(Path.Combine(reportsSaveLocation, reportsFolderName));

            return true;
        }

        /// <summary>
        /// Validate the input dates to make sure:
        /// - The dateFrom must be before dateTo
        /// - date range shall not exceed 7 days, if exceeded, show a message with user confirmation
        /// </summary>
        /// <param name="dateFrom">dateFrom</param>
        /// <param name="dateTo">dateTo</param>
        /// <returns>dates are all validated</returns>
        private bool validateRequestedDates(DateTime dateFrom, DateTime dateTo)
        {
            if (dateTo < dateFrom || dateTo > DateTime.Today || dateFrom > DateTime.Today)
            {
                MessageBox.Show(this, "Report daterange invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else if (dateTo - dateFrom > TimeSpan.FromDays(7))
            {
                DialogResult dlgRes = MessageBox.Show(this,
                    @"Date range exceed 7 days, the application will run very slow and may cause data loss. Continue?",
                    @"Continue", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dlgRes == DialogResult.No)
                    return false;
            }

            return true;
        }

        private void cbSelectReportOptionId_SelectedIndexChanged(object sender, EventArgs e)
        {
            var optionId = this.cbSelectReportOptionId.SelectedValue.ToString();
            if (optionId == NewReportOptionLabel)
            {
                this.CurrentReportOption = new ReportOption();
            }
            else
            {
                this.CurrentReportOption = Helpers.LoadReportOption(optionId);
                this.dirty = false;
            }
        }

        private void dtpFrom_ValueChanged(object sender, EventArgs e)
        {
            this.dtpEnd.Value = dtpStart.Value + TimeSpan.FromDays(7);
        }

        internal void frmExportConfig_Load(object sender, EventArgs e)
        {
            // load data into combo box
            this.LoadReportOptionNamesToComboBox();

            // load current report option from appData
            if (Settings.Default.LastReportOptionId == "") return;

            var _reportOption = Helpers.LoadReportOption(Settings.Default.LastReportOptionId);
            if (_reportOption != null)
            {
                this.CurrentReportOption = _reportOption; // load data into grid view
            }
            else
            {
                this.CurrentReportOption = new ReportOption();
            }

            // fill data into dgParts
            this.dgParts.AutoGenerateColumns = true;
            this.dgParts.DataSource = this.dgBindingList;

            foreach (var part in this.CurrentReportOption.Parts)
                this.dgBindingList.Add(part);
        }

        private void frmExportConfig_ReportOptionChanged(object sender, EventArgs e)
        {
            var _option = (ReportOption)sender;

            if (_option == null) return;

            this.cbSelectReportOptionId.SelectedItem = _option.Id;
            this.txtCustomerName.Text = _option.CustomerName;
            this.txtProcedure.Text = _option.Procedure;
            this.txtAcceptanceCriteria.Text = _option.AcceptanceCriteria;

            this.dgParts.AutoGenerateColumns = true;

            this.dgParts.DataSource = _option.Parts;
        }

        private void LoadReportOptionNamesToComboBox()
        {
            // load id into combobox
            this.cbSelectReportOptionId.Items.Clear();
            var availableOptions = Helpers.FetchReportOptions();
            string[] displayOptions;
            if (!availableOptions.Contains(NewReportOptionLabel))
            {
                displayOptions = new string[availableOptions.Length + 1];
                new string[] { NewReportOptionLabel }.CopyTo(displayOptions, 0);
                availableOptions.CopyTo(displayOptions, 1);
            }
            else
            {
                displayOptions = availableOptions;
            }

            this.cbSelectReportOptionId.DataSource = displayOptions;

            if (this.CurrentReportOption.Id != "")
                this.cbSelectReportOptionId.SelectedItem = this.CurrentReportOption.Id;
            else
                this.cbSelectReportOptionId.SelectedItem = NewReportOptionLabel;
        }

        internal IWorkbook MakeDailyReports(ReportKey[] reportKeys, DataTable sourceReport, string reportForm = "201")
        {
            // check if sourceReport contains no data, return null
            if (sourceReport == null || sourceReport.Rows.Count == 0)
                return null;

            // Working with mt1Report

            string templateName = Helpers.GetReportTemplate(reportForm);
            using (var mt1Template = new FileStream(templateName, FileMode.Open, FileAccess.Read))
            {
                var workbook = WorkbookFactory.Create(mt1Template);

                foreach (var reportKey in reportKeys)
                {
                    //get partInfo
                    var part = this.CurrentReportOption.Parts.FirstOrDefault(p => p.PartNo == reportKey.PartNo);
                    var lstExtendOfTest = new string[] { };
                    if (part != null)
                        lstExtendOfTest =
                            part.ExtendOfTest.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                    if (lstExtendOfTest.Length == 0)
                        lstExtendOfTest = new string[] { string.Empty };

                    var reportData = lstExtendOfTest.Join(sourceReport.AsEnumerable(), e => true, d => true,
                            (ExtendOfTest, reportRow) => new ReportDataRow()
                            {
                                ExtendOfTest = ExtendOfTest,
                                Type = reportRow.Field<string>("Type"),
                                SN = reportRow.Field<string>("SN"),
                                InspectedDate = reportRow.Field<DateTime>("Inspected Date"),
                                TestTemp = reportRow.Field<int>("Test Temp."),
                                Concentration = reportRow.Field<double>("Concentration"),
                                UV = reportRow.Field<int>("UV"),
                                VisibleLight = reportRow.Field<double>("Visible light"),
                                Shift = reportRow.Field<string>("Shift"),
                                Inspector = reportRow.Field<string>("Inspector"),
                                PartNo = reportRow.Field<string>("Part No."),
                            }).Where(r =>
                            r.InspectedDate == reportKey.InspectedDate &&
                            string.Equals(r.PartNo.Trim(), reportKey.PartNo.Trim(), StringComparison.CurrentCultureIgnoreCase) &&
                            r.Concentration == reportKey.Concentration &&
                            r.Shift.ToLower().Trim() == reportKey.Shift.ToLower().Trim())
                        .OrderBy(r => r.SN)
                        .ToList();

                    // paginate data into smaller chunks, each chunk contents are fit into a page. Each chunk is driven by the PartSN
                    var lstPaginatedDataDict = new List<Dictionary<string, IList<ReportDataRow>>>();
                    var dicDataOfCurPage = new Dictionary<string, IList<ReportDataRow>>();
                    var numOfRowFeededToCurPage = 0;
                    while (reportData.Count > 0 && numOfRowFeededToCurPage < OutReport.PAGE_SIZE)
                    {
                        // get the first SN from dataReport
                        var firstSerialNumber = reportData[0].SN;

                        // query number of ReportDataRow are being contained in reportData which has the same SN
                        var subCollectionOfReportData = reportData.Where(x => x.SN == firstSerialNumber).ToList();
                        var numOfReportDataRow = subCollectionOfReportData.Count;

                        // if sum of dataRowFeedTopage and numOfReportDataRow is less than PAGE_SIZE, then meant that the dictReportDataInPage still being available to add more data. Then add it into page feed, and remove from the whole reportData
                        if (numOfRowFeededToCurPage + numOfReportDataRow < OutReport.PAGE_SIZE)
                        {
                            dicDataOfCurPage.Add(firstSerialNumber, subCollectionOfReportData);
                            numOfRowFeededToCurPage += numOfReportDataRow;

                            // remove from dictReportData
                            subCollectionOfReportData.ForEach(x => reportData.Remove(x));
                        }
                        else
                        {
                            // put page into whole stack
                            lstPaginatedDataDict.Add(dicDataOfCurPage);

                            // no more adding, reset the dataRowFeedToPage to 0, reset the pageDict
                            numOfRowFeededToCurPage = 0;
                            dicDataOfCurPage = new Dictionary<string, IList<ReportDataRow>>();
                        }

                        // finish it, with a finish added
                        if (dicDataOfCurPage.Count != 0 && reportData.Count == 0)
                        {
                            // put page into whole stack
                            lstPaginatedDataDict.Add(dicDataOfCurPage);
                        }
                    }

                    // generate the report
                    var curPage = 1;
                    foreach (IDictionary<string, IList<ReportDataRow>> dictPageData in lstPaginatedDataDict)
                    {
                        this.FillDataToOutReportWorkbook(ref workbook, reportKey, dictPageData, part, curPage,
                            lstPaginatedDataDict.Count);
                        curPage++;
                    }
                }

                mt1Template.Close();

                return workbook;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="reportKey"></param>
        /// <param name="dictPageData"></param>
        /// <param name="part"></param>
        /// <param name="page"></param>
        /// <param name="totalPage"></param>
        private void FillDataToOutReportWorkbook(ref IWorkbook workbook, ReportKey reportKey,
            IDictionary<string, IList<ReportDataRow>> dictPageData, Part part, int page, int totalPage)
        {
            var sampleSheet = workbook.GetSheet("Report");
            var newSheetName = Helpers.makeReportSheetName(reportKey, page);

            // copy sample sheet into newSheetName, which is a page of the report
            sampleSheet.CopyTo(workbook, newSheetName, true, true);
            var newSheet = workbook.GetSheet(newSheetName);
            newSheet.PrintSetup.PaperSize = (short)PaperSize.A4 + 1;

            if (newSheet == null)
            {
                throw new CreateNewSheetException("Cannot create new sheet on output report");
            }

            var mt1 = new OutReport(newSheet);

            // general data
            var testTemp = 0;
            var uv = 0.0f;
            var visibleLight = 0.0;

            // fill report line
            var row = 0;
            foreach (var sn in dictPageData.Keys)
            {
                var lstReportData = dictPageData[sn];
                foreach (var reportRow in lstReportData)
                {
                    // extract testTemp, uv and visibleLight
                    testTemp = reportRow.TestTemp;
                    uv = reportRow.UV;
                    visibleLight = reportRow.VisibleLight;

                    mt1.SetLineValue(row, reportRow.SN, reportRow.ExtendOfTest);
                    row++;
                }

                // if lstReportData is not a single row, then need to merge the SN of these line
                if (lstReportData.Count <= 1) continue;

                try
                {
                    mt1.MergeLineSN(row - lstReportData.Count, row - 1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            mt1.SetFieldValue(CellLabels.REPORT_NO, $"VN-CQ-I19B1.7-{reportKey.InspectedDate:yyMM}-"); // this is MT1
            mt1.SetFieldValue(CellLabels.PART_NO, reportKey.PartNo);
            mt1.SetFieldValue(CellLabels.PART_NAME, part != null ? part.Description : string.Empty);
            mt1.SetFieldValue(CellLabels.CUSTOMER_NAME, this.CurrentReportOption.CustomerName);
            mt1.SetFieldValue(CellLabels.PROCEDURE, this.CurrentReportOption.Procedure);
            mt1.SetFieldValue(CellLabels.ACCEPTANCE_CRITERIA, this.CurrentReportOption.AcceptanceCriteria);
            mt1.SetFieldValue(CellLabels.CONCENTRATION, reportKey.Concentration);
            mt1.SetFieldValue(CellLabels.TEST_TEMP, testTemp);
            mt1.SetFieldValue(CellLabels.COIL, part?.Coil ?? 0);
            mt1.SetFieldValue(CellLabels.YOKE, part?.Yoke ?? 0);

            mt1.SetFieldValue(CellLabels.UV, uv);

            mt1.SetFieldValue(CellLabels.VISIBLE_LIGHT_INTENSITY, visibleLight);
            mt1.SetFieldValue(CellLabels.DATE_OF_EXAMINATION, reportKey.InspectedDate);
            mt1.SetFieldValue(CellLabels.DATE_OF_REPORT, reportKey.InspectedDate);
            mt1.SetFieldValue(CellLabels.PAGE_OF_TOTAL, $"Page {page}/{totalPage}");

            mt1.SetFieldValue(CellLabels.INSPECTOR, reportKey.Inspector);
        }

        private static void SaveReportFile(IWorkbook workbook, string savePath, DateTime date,
            string reportForm = "201")
        {
            // if workbook is null, return
            if (workbook == null) return;

            var outFilePrefix = reportForm == "201" ? "MT201_" : "MT202_";
            using (var mt1Report =
                   new FileStream(Path.Combine(savePath, string.Format("{1}{0:yyMMdd}.xlsx", date, outFilePrefix)),
                       FileMode.OpenOrCreate, FileAccess.Write))
            {
                if (workbook.NumberOfSheets > 1)
                    workbook.RemoveSheetAt(workbook.GetSheetIndex("Report"));
                workbook.Write(mt1Report);
                mt1Report.Close();
            }
        }

        private void txtAcceptanceCriteria_Validating(object sender, CancelEventArgs e)
        {
            this.ValidateTextFields(txtAcceptanceCriteria);
        }

        private void txtCustomerName_Validating(object sender, CancelEventArgs e)
        {
            this.ValidateTextFields(txtCustomerName);
        }

        private void txtProcedure_Validating(object sender, CancelEventArgs e)
        {
            this.ValidateTextFields(txtProcedure);
        }

        private void ValidateTextFields(Control sender)
        {
            if (sender.Text.Length == 0)
            {
                this.validated.Add(false);
                sender.BackColor = Color.Red;
            }
            else
            {
                this.dirty = true;
                sender.BackColor = SystemColors.Window;
            }
        }

        private void mt201_CheckedChanged(object sender, EventArgs e)
        {
            this.mt202.Checked = !this.mt201.Checked;
        }

        private void mt202_CheckedChanged(object sender, EventArgs e)
        {
            this.mt201.Checked = !this.mt202.Checked;
        }

        private void addNewPartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.dgBindingList.AddNew();
        }
    }
}