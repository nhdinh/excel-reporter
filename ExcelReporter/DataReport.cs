using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReporter
{
    public struct CellValue
    {
        public object Value;
        public Type ValueType;
    }

    public struct HeaderField
    {
        public Type ColumnType;
        public string HeaderLabel;

        public HeaderField(string HeaderLabel, Type ColumnType)
        {
            this.HeaderLabel = HeaderLabel;
            this.ColumnType = ColumnType;
        }
    }

    public class ReportRow
    {
        public string Type { get; set; }
        public string SN { get; set; }
        public string ReportNo { get; set; }
        public DateTime InspectedDate { get; set; }
        public int TestTemp { get; set; }
        public double Concentration { get; set; }
        public int UV { get; set; }
        public double VisibleLight { get; set; }
        public string AssyNo { get; set; }
        public string Remark { get; set; }
        public string Shift { get; set; }
        public string Inspector { get; set; }
        public string PartNo { get; set; }
    }

    public enum LoadStatus
    {
        LOADED = 0,
        UNLOADED = 1
    }

    public class DataReport
    {
        private string filePath = null;
        private HeaderField[] headerDatas = null;
        private int headerRowIndex = -1;
        private string sheetName = null;
        private string[] sheetNames = null;
        private IList<ReportRow> reportData = null;
        private WorkFileInfo workFileInfo;
        private LoadStatus loadStatus = ExcelReporter.LoadStatus.UNLOADED;
        private int totalItems = 0;

        public string FilePath
        {
            get
            {
                return filePath;
            }
        }

        public HeaderField[] HeaderDatas
        {
            get
            {
                return headerDatas;
            }
        }

        public int HeaderRowIndex
        {
            get
            {
                return headerRowIndex;
            }
        }

        public Guid Id
        {
            get
            {
                if (this.filePath != null)
                    using (MD5 md5 = MD5.Create())
                    {
                        byte[] hash = md5.ComputeHash(Encoding.Default.GetBytes(this.filePath.ToLower()));
                        Guid result = new Guid(hash);

                        return result;
                    }
                else
                    throw new Exception("FilePath not set");
            }
        }

        public LoadStatus LoadStatus
        {
            get
            {
                return this.loadStatus;
            }
        }

        public string SheetName
        {
            get
            {
                return sheetName;
            }
        }

        public string[] SheetNames
        {
            get
            {
                return this.sheetNames;
            }

            internal set
            {
                this.sheetNames = value;
            }
        }

        public WorkFileInfo WorkFile
        {
            get
            {
                return this.workFileInfo;
            }

            set
            {
                this.workFileInfo = value;
            }
        }

        public string Tag
        {
            get
            {
                return this.workFileInfo.Id.ToString();
            }
        }

        public int TotalItems
        {
            get
            {
                return this.totalItems;
            }

            internal set
            {
                this.totalItems = value;
            }
        }

        public IList<ReportRow> ReportData
        {
            get { return this.reportData; }
            internal set { this.reportData = value; }
        }

        /// <summary>
        /// Return names of workbook's sheets, provided its file path
        /// </summary>
        /// <param name="path">Path of the workbook file</param>
        /// <returns>Array of names</returns>
        public static string[] GetSheetNames(string path)
        {
            IWorkbook wb = new XSSFWorkbook(File.OpenRead(path));
            return GetSheetNames(wb);
        }

        public static DataReport LoadFile(string path, string sheetName = "", int headerRowIndex = 0, HeaderField[] headerFields = null)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IWorkbook wb = WorkbookFactory.Create(fs);

                // check for the sheet name
                string[] sheetNames = DataReport.GetSheetNames(wb);

                if (sheetName != "")
                {
                    if (!sheetNames.Contains(sheetName))
                    {
                        throw new SheetNotFoundException(path, sheetName);
                    }
                }
                else
                {
                    if (sheetNames.Length == 1)
                        sheetName = sheetNames[0];
                }

                // make new DataReport
                DataReport report = new DataReport
                {
                    sheetName = sheetName,
                    filePath = path,
                    sheetNames = sheetNames,
                    reportData = new List<ReportRow>()
                };

                // set headerRowIndex and convert data from sheet to List<ReportRow>
                ISheet workSheet = wb.GetSheet(sheetName);
                if (workSheet != null)
                {
                    if (headerRowIndex == 0)
                        report.headerRowIndex = report.DetermineHeaderRowIndexByHeaders(workSheet, headerFields);
                    else
                        report.headerRowIndex = headerRowIndex;

                    // if report.headerDatas is null, then collectHeaderLabels from worksheet
                    if (report.headerDatas == null)
                        report.headerDatas = report.CollectHeaderLabels(workSheet, report.headerRowIndex);

                    // load data from workSheet to List<ReportRow>
                    IRow sheetRow = null;
                    IList<string[]> stringRows = new List<string[]>();
                    for (int i = report.headerRowIndex + 1; i < workSheet.PhysicalNumberOfRows; i++)
                    {
                        sheetRow = workSheet.GetRow(i);
                        if (sheetRow != null && sheetRow.PhysicalNumberOfCells >= report.HeaderDatas.Length)
                        {
                            var stringRow = new string[13];
                            for (int j = 0; j < 13; j++)
                            {
                                stringRow[j] = GetCellValue(sheetRow.Cells[j]).ToString();
                            }

                            stringRows.Add(stringRow);
                        }
                    }

                    // clean the stringRows
                    var stringRowsArr = stringRows.Where(row => (
                    row[1] != String.Empty &&
                    row[3] != String.Empty &&
                    row[4] != String.Empty &&
                    row[5] != String.Empty &&
                    row[6] != String.Empty &&
                    row[7] != String.Empty &&
                    row[10] != String.Empty &&
                    row[11] != String.Empty &&
                    row[12] != String.Empty
                    )).ToArray();

                    for (int i = 0; i < stringRowsArr.Length; i++)
                    {
                        try
                        {
                            var reportRow = new ReportRow
                            {
                                Type = stringRowsArr[i][0],
                                SN = stringRowsArr[i][1],
                                ReportNo = stringRowsArr[i][2],

                                InspectedDate = DateTime.Parse(stringRowsArr[i][3]),
                                TestTemp = int.Parse(stringRowsArr[i][4]),
                                Concentration = double.Parse(stringRowsArr[i][5]),

                                UV = int.Parse(stringRowsArr[i][6]),
                                VisibleLight = double.Parse(stringRowsArr[i][7]),
                                AssyNo = stringRowsArr[i][8],
                                Remark = stringRowsArr[i][9],
                                Shift = stringRowsArr[i][10],
                                Inspector = stringRowsArr[i][11],
                                PartNo = stringRowsArr[i][12]
                            };

                            //  add to report instance
                            report.reportData.Add(reportRow);
                        }
                        catch
                        {
                            Console.WriteLine("Ignored row #{0}", i);
                        }
                    };

                    report.totalItems = report.reportData.Count;
                    report.loadStatus = LoadStatus.LOADED;
                }

                stopwatch.Stop();
                TimeSpan ts = stopwatch.Elapsed;
                string elapsedTime = String.Format("Load excel file: {0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                Console.WriteLine(elapsedTime);

                return report;
            }
        }

        public WorkFileInfo GenerateWorkFile()
        {
            var workFile = new WorkFileInfo
            {
                FilePath = this.filePath,
                SheetName = this.sheetName,
                HeaderRowIndex = this.headerRowIndex,
                HeaderLabels = this.headerDatas
            };

            return workFile;
        }

        internal HeaderField[] CollectHeaderLabels(ISheet workSheet, int headerRowIndex = -1)
        {
            if (headerRowIndex == -1)
                headerRowIndex = this.headerRowIndex;

            if (workSheet != null)
            {
                var headerRow = workSheet.GetRow(headerRowIndex);
                if (headerRow != null)
                {
                    int cellCount = headerRow.PhysicalNumberOfCells;
                    var _headerDatas = new HeaderField[cellCount];

                    Parallel.For(0, cellCount - 1, index =>
                      {
                          var cellVal = GetCellValue(headerRow.Cells[index]);

                          if (!cellVal.ToString().StartsWith("Column"))
                              _headerDatas[index] = new HeaderField(cellVal.ToString(), typeof(object));
                      });

                    return _headerDatas.Where(x => !string.IsNullOrEmpty(x.HeaderLabel)).ToArray();
                }
                else
                    throw new HeaderRowNotFoundException();
            }

            return new HeaderField[0];
        }

        /// <summary>
        /// Determine the header row of a working sheet by matching the label and row content
        /// </summary>
        /// <returns>Indexing number of the header row</returns>
        internal int DetermineHeaderRowIndexByHeaders(ISheet workSheet, HeaderField[] headerFields)
        {
            int rowToScan = 20;
            double maxMatchedPoint = 0.0;
            int maxMatchedRowIndex = 0;
            string[] headerLabels = headerFields.Select(x => x.HeaderLabel).ToArray();
            int unmatchedLabel = headerFields.Length;

            if (workSheet == null)
                throw new Exception("Worksheet is null");

            if (workSheet != null)
            {
                for (int i = 0; i <= rowToScan; i++)
                {
                    var row = workSheet.GetRow(i);
                    double matchedPoint = 0.0;

                    if (row.PhysicalNumberOfCells >= headerLabels.Length)
                        for (int j = 0; j < headerLabels.Length; j++)
                        {
                            var cellValue = GetCellValue(row.Cells[j]);
                            if (headerLabels.Contains(cellValue))
                            {
                                matchedPoint += 1.0 / headerLabels.Length;
                                unmatchedLabel--;
                            }
                        }

                    if (matchedPoint > maxMatchedPoint)
                    {
                        maxMatchedPoint = matchedPoint;
                        maxMatchedRowIndex = i;
                    }
                }

                // If there is no unmatched label, then copy the default header datas to report
                if (unmatchedLabel == 0)
                {
                    this.headerDatas = (HeaderField[])headerFields.Clone();
                }

                return maxMatchedRowIndex;
            }

            return -1;
        }

        internal List<ReportRow> GetData(DateTime fromDate, DateTime toDate)
        {
            return this.reportData.Where(row => row.InspectedDate >= fromDate && row.InspectedDate <= toDate).ToList();
        }

        internal DataTable GetDataAsDataTable(DateTime fromDate, DateTime toDate)
        {
            List<ReportRow> returnList = this.GetData(fromDate, toDate);
            DataTable returnTable = new DataTable();

            // handle columns
            foreach (var headerData in headerDatas)
                returnTable.Columns.Add(headerData.HeaderLabel, headerData.ColumnType);

            foreach (var reportRow in returnList)
                returnTable.Rows.Add(
                    reportRow.Type,
                    reportRow.SN,
                    reportRow.ReportNo,
                    reportRow.InspectedDate,
                    reportRow.TestTemp,
                    reportRow.Concentration,
                    reportRow.UV,
                    reportRow.VisibleLight,
                    reportRow.AssyNo,
                    reportRow.Remark,
                    reportRow.Shift,
                    reportRow.Inspector,
                    reportRow.PartNo
                    );

            return returnTable;
        }

        /// <summary>
        /// Return names of workbook's sheets
        /// </summary>
        /// <param name="wb">Instance of the workbook</param>
        /// <returns>Array of names</returns>
        private static string[] GetSheetNames(IWorkbook wb)
        {
            int sheetCount = wb.NumberOfSheets;
            string[] sheetNames = new string[sheetCount];

            for (int i = 0; i < sheetCount; i++)
            {
                sheetNames[i] = wb.GetSheetName(i);
            }

            return sheetNames;
        }

        private static object GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    else
                    {
                        return cell.NumericCellValue;
                    }
                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Formula:
                    return cell.CellFormula;

                case CellType.Blank:
                    return String.Empty;

                default:
                    return cell.CellType.ToString();
            }
        }

        internal IList<ReportRow> GetReportData(int firstIndex, int lastIndex)
        {
            return null;
        }
    }
}