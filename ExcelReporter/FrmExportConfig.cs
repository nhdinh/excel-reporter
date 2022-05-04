using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static System.Environment;

namespace ExcelReporter
{
    public partial class FrmExportConfig : Form
    {
        private readonly CoreApp app;
        private readonly string newReportOptionLabel = "New...";
        private bool dirty = false;
        private ReportOption reportOption;
        private List<bool> validated = new List<bool>();
        private BindingList<Part> dgBindingList = new BindingList<Part>();

        public FrmExportConfig()
        {
            InitializeComponent();

            // init app
            this.app = CoreApp.GetInstance();
            this.ReportOptionChanged += FrmExportConfig_ReportOptionChanged;
        }

        public event EventHandler ReportOptionChanged;

        protected ReportOption ReportOption
        {
            get { return this.reportOption; }
            private set
            {
                this.reportOption = value;
                this.OnReportOptionChanged(this.reportOption);
            }
        }

        protected virtual void OnReportOptionChanged(ReportOption option)
        {
            ReportOptionChanged?.Invoke(option, EventArgs.Empty);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            if (dirty)
            {
                DialogResult res = MessageBox.Show(this, "Data is not saved. Close anyway?", "Close?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

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

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            string reportOptionId = (string)this.cbSelectReportOptionId.SelectedItem;
            if (reportOptionId != this.newReportOptionLabel)
                try
                {
                    Helpers.DeleteReportOption(reportOptionId);
                    this.LoadReportOptionNamesToComboBox();
                }
                catch
                {
                    MessageBox.Show(this, "Delete failed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
        }

        private void BtnSaveAndClose_Click(object sender, EventArgs e)
        {
            this.validated = new List<bool>();

            foreach (var textbox in groupBox1.Controls)
            {
                if (textbox.GetType() == typeof(TextBox))
                {
                    this.ValidateTextFields((TextBox)textbox, null);
                }
            }

            if (!this.validated.Where(x => x == false).Any())
            {
                string reportOptionId = (string)this.cbSelectReportOptionId.SelectedItem;
                if (reportOptionId == this.newReportOptionLabel)
                {
                    this.reportOption.Id = Helpers.MakeNewReportOptionId(this.txtCustomerName.Text);
                }

                this.reportOption.CustomerName = this.txtCustomerName.Text.Trim();
                this.reportOption.Procedure = this.txtProcedure.Text.Trim();
                this.reportOption.AcceptanceCriteria = this.txtAcceptanceCriteria.Text.Trim();

                this.reportOption.Parts = new List<Part>();
                foreach (DataGridViewRow row in this.dgParts.Rows)
                {
                    if (row.Cells["partno"].Value != null)
                    {
                        var part = new Part()
                        {
                            PartNo = row.Cells["PartNo"].Value != null ? (string)row.Cells["PartNo"].Value : "",
                            Description = row.Cells["Description"].Value != null ? (string)row.Cells["Description"].Value : "",
                            Coil = int.Parse(row.Cells["Coil"].Value.ToString()),
                            Yoke = int.Parse(row.Cells["Yoke"].Value.ToString()),
                            ExtendOfTest = row.Cells["ExtendOfTest"].Value != null ? (string)row.Cells["ExtendOfTest"].Value : "",
                        };
                        this.reportOption.Parts.Add(part);
                    }
                }

                // save report option
                Helpers.SaveReportOption(this.reportOption);

                // save config
                this.app.Config.LastReportOptionId = this.reportOption.Id;
                Helpers.SaveAppConfig(this.app.Config);

                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void btnSaveAndExport_Click(object sender, EventArgs e)
        {
            var dateFrom = this.dtpFrom.Value.Date;
            var dateTo = this.dtpTo.Value.Date;
            string reportForm = this.mt201.Checked ? "201" : "202";

            if (dateTo < dateFrom || dateTo > DateTime.Today || dateFrom > DateTime.Today)
            {
                MessageBox.Show(this, "Report daterange invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (dateTo - dateFrom > TimeSpan.FromDays(7))
            {
                DialogResult dlgRes = MessageBox.Show(this, "Date range exceed 7 days, the application will run very slow and may cause data loss. Continue?", "Continue", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dlgRes == DialogResult.No)
                    return;
            }

            // make data
            DataSet dataset = null;
            try
            {
                dataset = this.app.MakeReportData(dateFrom, dateTo);
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, "Error while making report data. " + exc.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }

            if (dataset != null)
            {
                // create save location in My Documents folder
                string reportsSaveLocation = Path.Combine(GetFolderPath(SpecialFolder.MyDocuments), "Reports");
                string reportsFolderName = String.Format("{0:yyMMdd}_{1:yyMMdd}", dateFrom, dateTo);

                // check the target save folder if existed and contains files, then need asking to backup, else create new
                if (Directory.Exists(Path.Combine(reportsSaveLocation, reportsFolderName)))
                {
                    if (Directory.GetFiles(Path.Combine(reportsSaveLocation, reportsFolderName), "*.*").Length != 0)
                    {
                        DialogResult res = MessageBox.Show(this, string.Format("Report for {0} already exists. Make backup and continue?", reportsFolderName), "Reports exists", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                        if (res == DialogResult.Yes)
                        {
                            try
                            {                             // move to another folder
                                Directory.Move(Path.Combine(reportsSaveLocation, reportsFolderName), Path.Combine(reportsSaveLocation, string.Format("{0}_backup{1:yyMMddHHmmss}", reportsFolderName, DateTime.Now)));

                                // recreate the report folder
                                Directory.CreateDirectory(Path.Combine(reportsSaveLocation, reportsFolderName));
                            }
                            catch
                            {
                                MessageBox.Show("There is an error while creating the backup, close any opening related files and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        else
                        {
                            return; // cancel
                        }
                    }
                }
                else
                    // create directory
                    Directory.CreateDirectory(Path.Combine(reportsSaveLocation, reportsFolderName));
                DateTime[] reportDateRange = dataset.Tables["report_key"].AsEnumerable()
                    .GroupBy(g => new { InspectedDate = g.Field<DateTime>("Inspected Date") })
                    .Where(g => g.Count() >= 1)
                    .OrderBy(g => g.Key.InspectedDate)
                    .Select(g => new DateTime(
                        g.Key.InspectedDate.Year,
                        g.Key.InspectedDate.Month,
                        g.Key.InspectedDate.Day
                        )).ToArray();

                // loop through each date in the reportDateRange and create workbook for each date
                foreach (var reportDate in reportDateRange)
                {
                    DateTime dateValue = reportDate;

                    ReportKey[] reportKeys = dataset.Tables["report_key"].AsEnumerable()
                        .GroupBy(row => new
                        {
                            InspectedDate = row.Field<DateTime>("Inspected Date"),
                            PartNo = row.Field<string>("Part No."),
                            Concentration = row.Field<double>("Concentration"),
                            Inspector = row.Field<string>("Inspector"),
                            Shift = row.Field<string>("Shift")
                        }).Where(g => g.Key.InspectedDate == dateValue)
                    .Select(o => new ReportKey(
                        o.Key.InspectedDate,
                        o.Key.PartNo,
                        o.Key.Concentration,
                        o.Key.Inspector,
                        o.Key.Shift
                        )).ToArray();

                    this.MakeDailyReports(Path.Combine(reportsSaveLocation, reportsFolderName), dateValue, reportKeys, dataset.Tables["source"], reportForm);
                };

                MessageBox.Show(this, string.Format("Exported successfully to {0}.", reportsSaveLocation), "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void CbSelectReportOptionId_SelectedIndexChanged(object sender, EventArgs e)
        {
            var optionId = this.cbSelectReportOptionId.SelectedValue.ToString();
            if (optionId == this.newReportOptionLabel)
            {
                this.ReportOption = new ReportOption();
            }
            else
            {
                this.ReportOption = Helpers.LoadReportOption(optionId);
                this.dirty = false;
            }
        }

        private void DtpFrom_ValueChanged(object sender, EventArgs e)
        {
            this.dtpTo.Value = dtpFrom.Value + TimeSpan.FromDays(7);
        }

        private void FrmExportConfig_Load(object sender, EventArgs e)
        {
            // load data into combo box
            this.LoadReportOptionNamesToComboBox();

            // load current report option from appDaata
            if (this.app.Config.LastReportOptionId != "")
            {
                var _reportOption = Helpers.LoadReportOption(this.app.Config.LastReportOptionId);
                if (_reportOption != null)
                {
                    this.ReportOption = _reportOption;// load data into gridview
                    //this.dgParts.DataSource = this.makeDataTable(this.reportOption.Parts);

                    this.dgParts.AutoGenerateColumns = true;
                    //this.dgParts.DataSource = this.makeDataTable(_option.Parts);
                    this.dgParts.DataSource = this.dgBindingList;

                    foreach (var part in this.reportOption.Parts)
                        this.dgBindingList.Add(part);
                }
                else
                {
                    this.ReportOption = new ReportOption();
                }
            }
        }

        private void FrmExportConfig_ReportOptionChanged(object sender, EventArgs e)
        {
            var _option = (ReportOption)sender;

            if (_option != null)
            {
                this.cbSelectReportOptionId.SelectedItem = _option.Id;
                this.txtCustomerName.Text = _option.CustomerName;
                this.txtProcedure.Text = _option.Procedure;
                this.txtAcceptanceCriteria.Text = _option.AcceptanceCriteria;

                //this.dgParts.Rows.Clear();
                this.dgParts.AutoGenerateColumns = true;
                //this.dgParts.DataSource = this.makeDataTable(_option.Parts);
                this.dgParts.DataSource = _option.Parts;
            }
        }

        private void LoadReportOptionNamesToComboBox()
        {
            // load id into combobox
            this.cbSelectReportOptionId.Items.Clear();
            string[] availableOptions = Helpers.FetchReportOptions();
            string[] displayOptions;
            if (!availableOptions.Contains(this.newReportOptionLabel))
            {
                displayOptions = new string[availableOptions.Length + 1];
                new string[] { this.newReportOptionLabel }.CopyTo(displayOptions, 0);
                availableOptions.CopyTo(displayOptions, 1);
            }
            else
            {
                displayOptions = availableOptions;
            }

            this.cbSelectReportOptionId.DataSource = displayOptions;

            if (this.reportOption.Id != "")
                this.cbSelectReportOptionId.SelectedItem = this.reportOption.Id;
            else this.cbSelectReportOptionId.SelectedItem = this.newReportOptionLabel;
        }

        private void MakeDailyReports(string savePlace, DateTime date, ReportKey[] reportKeys, DataTable sourceReport, string reportForm = "201")
        {
            // Working with mt1Report
            IWorkbook workbook;
            ISheet sampleSheet;

            // TODO: FIleStream will use more memory than File
            string templateName = reportForm == "201" ? @".\templates\mt201_template.xlsx" : @".\templates\mt202_template.xlsx";
            using (var mt1template = new FileStream(templateName, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(mt1template);
                sampleSheet = workbook.GetSheet("Report");

                foreach (ReportKey reportKey in reportKeys)
                {
                    //get partInfo
                    Part part = this.reportOption.Parts.FirstOrDefault(p => p.PartNo == reportKey.PartNo);
                    string[] lstExtendOfTest = new string[] { };
                    if (part != null)
                        lstExtendOfTest = part.ExtendOfTest.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

                    if (lstExtendOfTest.Length == 0)
                        lstExtendOfTest = new string[] { string.Empty };

                    var reportData = lstExtendOfTest.Join(sourceReport.AsEnumerable(), e => true, d => true, (ExtendOfTest, reportRow) => new
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
                    }).Where(r => r.InspectedDate == reportKey.InspectedDate && r.PartNo.Trim().ToLower() == reportKey.PartNo.Trim().ToLower() && r.Concentration == reportKey.Concentration && r.Shift.ToLower().Trim() == reportKey.Shift.ToLower().Trim())
                        .OrderBy(r => r.InspectedDate).OrderBy(r => r.SN).ToList();

                    // calculate number of page
                    int totalRows = reportData.Count;
                    int pageNum = (int)Math.Ceiling((double)totalRows / ExcelReport.PAGE_SIZE);

                    for (int page = 1; page <= pageNum; page++)
                    {
                        string newSheetName;
                        if (pageNum > 1)
                            newSheetName = String.Format("{0:dd}_{1}_{2}_{3}{4}_P{5}", reportKey.InspectedDate, reportKey.Concentration * 100, reportKey.PartNo, reportKey.Inspector.Remove(3), reportKey.Shift.Replace("S", ""), page);
                        else
                            newSheetName = String.Format(
                                "{0:dd}_{1}_{2}_{3}{4}",
                                reportKey.InspectedDate,
                                reportKey.Concentration * 100,
                                reportKey.PartNo,
                                reportKey.Inspector.Length > 3 ? reportKey.Inspector.Remove(3) : reportKey.Inspector,
                                reportKey.Shift.Replace("S", "")
                                );

                        sampleSheet.CopyTo(workbook, newSheetName, true, true);
                        ISheet newSheet = workbook.GetSheet(newSheetName);
                        newSheet.PrintSetup.PaperSize = (short)PaperSize.A4 + 1;

                        if (newSheet != null)
                        {
                            ExcelReport mt1 = new ExcelReport
                            {
                                Data = newSheet
                            };

                            // general data
                            int testTemp = 0;
                            double uv = 0.0;
                            double visibleLight = 0.0;

                            // fill data

                            int endRow = page * ExcelReport.PAGE_SIZE;
                            if (endRow > totalRows)
                                endRow = totalRows;
                            for (int row = (page - 1) * ExcelReport.PAGE_SIZE; row < endRow; row++)
                            {
                                if (totalRows > row)
                                {
                                    var reportRow = reportData[row];

                                    testTemp = reportRow.TestTemp;
                                    uv = reportRow.UV;
                                    visibleLight = reportRow.VisibleLight;

                                    mt1.SetLineValue(row, reportRow.SN, reportRow.ExtendOfTest);
                                }
                            }

                            mt1.SetFieldValue(CELL_ADDRESSES.REPORT_NO, string.Format("VN-CQ-I19B1.7-{0:yyMM}-", reportKey.InspectedDate)); // this is MT1
                            mt1.SetFieldValue(CELL_ADDRESSES.PART_NO, reportKey.PartNo);
                            mt1.SetFieldValue(CELL_ADDRESSES.PART_NAME, part != null ? part.Description : String.Empty);
                            mt1.SetFieldValue(CELL_ADDRESSES.CUSTOMER_NAME, this.reportOption.CustomerName);
                            mt1.SetFieldValue(CELL_ADDRESSES.PROCEDURE, this.reportOption.Procedure);
                            mt1.SetFieldValue(CELL_ADDRESSES.ACCEPTANCE_CRITERIA, this.reportOption.AcceptanceCriteria);
                            mt1.SetFieldValue(CELL_ADDRESSES.CONCENTRATION, reportKey.Concentration);
                            mt1.SetFieldValue(CELL_ADDRESSES.TEST_TEMP, testTemp);
                            mt1.SetFieldValue(CELL_ADDRESSES.COIL, part != null ? part.Coil : 0);
                            mt1.SetFieldValue(CELL_ADDRESSES.YOKE, part != null ? part.Yoke : 0);

                            mt1.SetFieldValue(CELL_ADDRESSES.UV, uv);

                            mt1.SetFieldValue(CELL_ADDRESSES.VISIBLE_LIGHT_INTENSITY, visibleLight);
                            mt1.SetFieldValue(CELL_ADDRESSES.DATE_OF_EXAMINATION, reportKey.InspectedDate);
                            mt1.SetFieldValue(CELL_ADDRESSES.DATE_OF_REPORT, reportKey.InspectedDate);
                            mt1.SetFieldValue(CELL_ADDRESSES.PAGE_OF_TOTAL, string.Format("Page {0}/{1}", page, pageNum));

                            mt1.SetFieldValue(CELL_ADDRESSES.INSPECTOR, reportKey.Inspector);
                        }
                    }
                }

                string outFilePrefix = reportForm == "201" ? "MT201_" : "MT202_";
                using (FileStream mt1report = new FileStream(Path.Combine(savePlace, String.Format("{1}{0:yyMMdd}.xlsx", date, outFilePrefix)), FileMode.OpenOrCreate, FileAccess.Write))
                {
                    if (workbook.NumberOfSheets > 1)
                        workbook.RemoveSheetAt(workbook.GetSheetIndex("Report"));
                    workbook.Write(mt1report);
                    mt1report.Close();
                }

                mt1template.Close();
            }
        }

        private void TxtAcceptanceCriteria_Validating(object sender, CancelEventArgs e)
        {
            this.ValidateTextFields(txtAcceptanceCriteria, e);
        }

        private void TxtCustomerName_Validating(object sender, CancelEventArgs e)
        {
            this.ValidateTextFields(txtCustomerName, e);
        }

        private void TxtProcedure_Validating(object sender, CancelEventArgs e)
        {
            this.ValidateTextFields(txtProcedure, e);
        }

        private void ValidateTextFields(TextBox sender, CancelEventArgs e)
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

        private void Mt201_CheckedChanged(object sender, EventArgs e)
        {
            this.mt202.Checked = !this.mt201.Checked;
        }

        private void Mt202_CheckedChanged(object sender, EventArgs e)
        {
            this.mt201.Checked = !this.mt202.Checked;
        }
    }
}