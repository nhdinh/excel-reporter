using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace ExcelReporter
{
    internal class CoreApp : INotifyPropertyChanged
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static CoreApp _instance = null;
        private AppConfig _appConfig;
        private IList<DataReport> _reports = null;
        //private IList<WorkFileInfo> _workFiles = null;

        private CoreApp()
        {
            this._appConfig = Helpers.LoadAppConfigFromJson();
            this._appConfig.DefaultHeaderFields = new HeaderField[] {
                    new HeaderField("Type", typeof(string)),
                    new HeaderField("SN", typeof(string)),
                    new HeaderField("Report No.", typeof(string)),
                    new HeaderField("Inspected Date", typeof(DateTime)),
                    new HeaderField("Test Temp.", typeof(int)),
                    new HeaderField("Concentration", typeof(double)),
                    new HeaderField("UV", typeof(int)),
                    new HeaderField("Visible light", typeof(double)),
                    new HeaderField("Ass'y No", typeof(string)),
                    new HeaderField("Remark", typeof(string)),
                    new HeaderField("Shift", typeof(string)),
                    new HeaderField("INSPECTOR", typeof(string)),
                    new HeaderField("Part No.", typeof(string))
                };
            this._reports = new List<DataReport>();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public AppConfig Config
        {
            get { return this._appConfig; }
            internal set { this._appConfig = value; }
        }

        public IList<DataReport> Reports
        {
            get
            {
                return _reports;
            }
        }

        public IList<WorkFileInfo> WorkFiles
        {
            get
            {
                //return _workFiles;
                return this._reports.Select(x => x.WorkFile).ToList();
            }

            private set
            {
                if (value != null)
                {
                    var loadedWorkFiles = this._reports.Select(r => r.WorkFile).ToList();
                    if (!loadedWorkFiles.Contains((WorkFileInfo)value))
                    {
                        NotifyPropertyChanged("");
                    }
                }
            }
        }

        public static CoreApp GetInstance()
        {
            if (_instance == null)
                _instance = new CoreApp();

            return _instance;
        }

        internal DataSet MakeReportData(DateTime dateFrom, DateTime dateTo)
        {
            DataSet returnDs = new DataSet();
            DataTable sourceTable = new DataTable("merged_source");
            foreach (DataReport report in this._reports)
            {
                var _table = report.GetDataAsDataTable(dateFrom, dateTo);
                sourceTable.Merge(_table);
            }

            var lstDuplicatedResult = sourceTable.AsEnumerable()
                .GroupBy(c => new
                {
                    Type = c.Field<string>("Type"),
                    SN = c.Field<string>("SN"),
                    ReportNo = c.Field<string>("Report No."),
                    InspectedDate = c.Field<DateTime>("Inspected Date"),
                    TestTemp = c.Field<int>("Test Temp."),
                    Concentration = c.Field<double>("Concentration"),
                    UV = c.Field<int>("UV"),
                    VisibleLight = c.Field<double>("Visible light"),
                    AssyNo = c.Field<string>("Ass'y No"),
                    Remark = c.Field<string>("Remark"),
                    Shift = c.Field<string>("Shift"),
                    Inspector = c.Field<string>("Inspector"),
                    PartNo = c.Field<string>("Part No.")
                }).Where(g => g.Count() >= 1 && g.Key.PartNo.Trim() != "")
                .Select(g => new
                {
                    g.Key.Type,
                    g.Key.SN,
                    g.Key.ReportNo,
                    g.Key.InspectedDate,
                    g.Key.TestTemp,
                    g.Key.Concentration,
                    g.Key.UV,
                    g.Key.VisibleLight,
                    g.Key.AssyNo,
                    g.Key.Remark,
                    g.Key.Shift,
                    g.Key.Inspector,
                    g.Key.PartNo
                }).ToList();

            DataTable filteredReportTable = new DataTable("source");
            foreach (var headerData in this.Config.DefaultHeaderFields)
                filteredReportTable.Columns.Add(headerData.HeaderLabel, headerData.ColumnType);
            for (int i = 0; i < lstDuplicatedResult.Count; i++)
            {
                var item = lstDuplicatedResult[i];
                filteredReportTable.Rows.Add(
                    item.Type,
                    item.SN,
                    item.ReportNo,
                    item.InspectedDate,
                    item.TestTemp,
                    item.Concentration,
                    item.UV,
                    item.VisibleLight,
                    item.AssyNo,
                    item.Remark,
                    item.Shift,
                    item.Inspector,
                    item.PartNo
                    );
            }

            // select InspectedDate from result table
            DataTable reportKeyTable = new DataTable("report_key");
            reportKeyTable.Columns.Add("Inspected Date", typeof(DateTime));
            reportKeyTable.Columns.Add("Part No.", typeof(string));
            reportKeyTable.Columns.Add("Concentration", typeof(double));
            reportKeyTable.Columns.Add("Inspector", typeof(string));
            reportKeyTable.Columns.Add("Shift", typeof(string));

            var lstInspectedDatesRslt = filteredReportTable.AsEnumerable()
                .GroupBy(row => new
                {
                    InspectedDate = row.Field<DateTime>("Inspected Date"),
                    PartNo = row.Field<string>("Part No."),
                    Concentration = row.Field<double>("Concentration"),
                    Inspector = row.Field<string>("Inspector"),
                    Shift = row.Field<string>("Shift")
                }).Where(g => g.Count() >= 1)
                .OrderBy(g => g.Key.InspectedDate)
                .Select(g => new ReportKey(
                    g.Key.InspectedDate,
                    g.Key.PartNo,
                    g.Key.Concentration,
                    g.Key.Inspector,
                g.Key.Shift
                )).ToList();

            //var lstInspectedDatesRslt = reportKeyResult.ToList();
            for (int i = 0; i < lstInspectedDatesRslt.Count; i++)
            {
                var item = lstInspectedDatesRslt[i];
                reportKeyTable.Rows.Add(
                    item.InspectedDate,
                    item.PartNo,
                    item.Concentration,
                    item.Inspector,
                    item.Shift
                    );
            }

            returnDs.Tables.Add(filteredReportTable);
            returnDs.Tables.Add(reportKeyTable);
            return returnDs;
        }

        internal void UpdateWorkfile(WorkFileInfo workFile)
        {
            var selectedReport = this._reports.Where(r => r.WorkFile.FilePath == workFile.FilePath).First();

            this._reports.Remove(selectedReport);

            var newReport = this.loadWorkFileToReport(workFile);
            this._reports.Add(newReport);
            NotifyPropertyChanged(newReport.Tag);
        }

        internal void LoadWorkFiles(IList<WorkFileInfo> workingFiles)
        {
            var _workingFiles = workingFiles.ToArray();
            this.LoadWorkFiles(_workingFiles);
        }

        internal void LoadWorkFiles(WorkFileInfo[] workFiles)
        {
            this._reports = new List<DataReport>();

            // try loading these workFiles
            if (workFiles != null)
            {
                bool reportChanged = false;

                for (int i = 0; i < workFiles.Length; i++)
                {
                    var workFile = workFiles[i];

                    // make task and run, then return the result
                    Task<DataReport> task = Task.Run(() =>
                    {
                        return this.loadWorkFileToReport(workFile);
                    });

                    if (task.Result != null)
                    {
                        if (this._reports == null)
                            this._reports = new List<DataReport>();

                        var createdReport = (DataReport)task.Result;
                        try
                        {
                            var reportHasExisted = this._reports.Where(r => r.FilePath == createdReport.FilePath).First();
                            if (reportHasExisted != null)
                                this._reports.Remove(reportHasExisted);
                        }
                        catch (InvalidOperationException)
                        { }

                        this._reports.Add(createdReport);

                        reportChanged = true;
                    }
                }

                if (reportChanged)
                    NotifyPropertyChanged("");
            }
        }

        private void addReport(DataReport report)
        {
            if (this._reports == null)
                this._reports = new List<DataReport>();

            this._reports.Add(report);
        }

        private void NotifyPropertyChanged([CallerMemberName] string changedReportTag = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(changedReportTag));
        }

        /// <summary>
        /// receive the workFile and make report from it.
        /// </summary>
        /// <param name="workFile"></param>
        /// <returns></returns>
        private DataReport loadWorkFileToReport(WorkFileInfo workFile)
        {
            var report = DataReport.LoadFile(workFile.FilePath, workFile.SheetName, 0, this.Config.DefaultHeaderFields);

            if (report != null)
            {
                workFile.FilePath = report.FilePath;
                workFile.SheetName = report.SheetName;
                workFile.SheetNames = report.SheetNames;
                workFile.HeaderRowIndex = report.HeaderRowIndex;
                workFile.HeaderLabels = report.HeaderDatas;

                report.WorkFile = workFile;
                return report;
            }

            // TODO: return an unloaded-report here
            return null;
        }
    }
}