using NLog;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReporter
{
    public struct ReportDataRow
    {
        public string ExtendOfTest { get; set; }
        public string Type { get; set; }
        public string SN { get; set; }
        public DateTime InspectedDate { get; set; }
        public int TestTemp { get; set; }
        public double Concentration { get; set; }
        public int UV { get; set; }
        public double VisibleLight { get; set; }
        public string Shift { get; set; }
        public string Inspector { get; set; }
        public string PartNo { get; set; }
    }

    public struct CellAddress
    {
        public int Row { get; private set; }
        public int Col { get; private set; }

        public CellAddress(int row, string col)
        {
            this.Row = row - 1;
            this.Col = GetColumnNumber(col) - 1;
        }

        private static int GetColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }
    }

    public struct ReportKey
    {
        public DateTime InspectedDate { get; private set; }
        public string PartNo { get; private set; }
        public double Concentration { get; private set; }
        public string Inspector { get; private set; }
        public string Shift { get; private set; }

        public ReportKey(DateTime inspectedDate, string partNo, double concentration, string inspector, string shift)
        {
            this.InspectedDate = inspectedDate;
            this.PartNo = partNo;
            this.Concentration = concentration;
            this.Inspector = inspector;
            this.Shift = shift;
        }
    }

    public class CellLabels
    {
        public static readonly CellAddress REPORT_NO = new CellAddress(3, "AL");

        public static readonly CellAddress CUSTOMER_NAME = new CellAddress(5, "A");
        public static readonly CellAddress PART_NAME = new CellAddress(5, "V");
        public static readonly CellAddress PART_NO = new CellAddress(5, "AL");
        public static readonly CellAddress PROCEDURE = new CellAddress(7, "V");
        public static readonly CellAddress ACCEPTANCE_CRITERIA = new CellAddress(7, "AL");
        public static readonly CellAddress TEST_TEMP = new CellAddress(9, "V");
        public static readonly CellAddress CONCENTRATION = new CellAddress(17, "H");
        public static readonly CellAddress COIL = new CellAddress(15, "AO");
        public static readonly CellAddress YOKE = new CellAddress(16, "AO");

        public static readonly CellAddress BLACK_LIGHT_MODEL = new CellAddress(15, "AA");
        public static readonly CellAddress BLACK_LIGHT_SN = new CellAddress(16, "X");
        public static readonly CellAddress UV = new CellAddress(17, "Z");

        public static readonly CellAddress VISIBLE_LIGHT_INTENSITY = new CellAddress(19, "F");

        public static readonly CellAddress DATE_OF_EXAMINATION = new CellAddress(45, "AO");
        public static readonly CellAddress DATE_OF_REPORT = new CellAddress(46, "AO");
        public static readonly CellAddress PAGE_OF_TOTAL = new CellAddress(47, "AY");

        public static readonly CellAddress INSPECTOR = new CellAddress(46, "K");

        private CellLabels()
        {
            // do nothing to protect the static readonly properties
        }
    }

    public class OutReport
    {
        private readonly Logger logger;
        private const int COL_PART_SN = 0;
        private const int COL_EXTEND_OF_TEST = 12;
        private const int COL_QUANTITY = 18;
        private const int COL_LOCATION = 21;
        private const int COL_LENGTH = 24;
        private const int COL_WIDTH = 27;
        private const int COL_DEPTH = 30;
        private const int COL_AREA = 33;
        private const int COL_TYPE = 37;
        private const int COL_ACCEPT = 40;

        private ISheet sheet { get; set; }

        public const int START_BODY_ROW = 24;
        public const int PAGE_SIZE = 18;

        public OutReport(ISheet newSheet)
        {
            this.sheet = newSheet;
            this.logger = LogManager.GetCurrentClassLogger();
        }

        public void SetFieldValue(CellAddress targetCell, object value)
        {
            if (this.sheet == null)
            {
                throw new Exception("Must set Sheet instance first");
            }

            ICell cell = this.sheet.GetRow(targetCell.Row).GetCell(targetCell.Col);

            if (value is string stringValue)
            {
                cell.SetCellValue(stringValue);
                cell.SetCellType(CellType.String);
            }
            else if (value is DateTime time)
            {
                cell.SetCellValue(time);
                cell.SetCellType(CellType.Numeric);
            }
            else if (value is int intNum)
            {
                cell.SetCellValue(intNum);
                cell.SetCellType(CellType.Numeric);
            }
            else if (value is double doubleNum)
            {
                cell.SetCellValue(doubleNum);
                cell.SetCellType(CellType.Numeric);
            }
        }

        public void SetCustomerName(string customerName)
        {
            this.SetFieldValue(CellLabels.CUSTOMER_NAME, customerName);
        }

        public void SetProcedure(string procedure)
        {
            this.SetFieldValue(CellLabels.PROCEDURE, procedure);
        }

        internal void SetLineValue(int row, string partSN, string extendOfTest)
        {
            if (row >= PAGE_SIZE)
                row -= PAGE_SIZE;

            try
            {
                IRow curRow = this.sheet.GetRow(OutReport.START_BODY_ROW + row);

                ICell snCell = curRow.GetCell(COL_PART_SN);
                if (snCell != null)
                    snCell.SetCellValue(partSN);

                ICell extendOfTestCell = curRow.GetCell(COL_EXTEND_OF_TEST);
                if (extendOfTestCell != null)
                    extendOfTestCell.SetCellValue(extendOfTest);

                ICell qtyCell = curRow.GetCell(COL_QUANTITY);
                if (qtyCell != null)
                    qtyCell.SetCellValue(1);

                ICell locationCell = curRow.GetCell(COL_LOCATION);
                if (locationCell != null)
                    locationCell.SetCellValue("–");

                ICell lengthCell = curRow.GetCell(COL_LENGTH);
                if (lengthCell != null)
                    lengthCell.SetCellValue("–");

                ICell widthCell = curRow.GetCell(COL_WIDTH);
                if (widthCell != null)
                    widthCell.SetCellValue("–");

                ICell depthCell = curRow.GetCell(COL_DEPTH);
                if (depthCell != null)
                    depthCell.SetCellValue("–");

                ICell areaCell = curRow.GetCell(COL_AREA);
                if (areaCell != null)
                    areaCell.SetCellValue("–");

                ICell typeCell = curRow.GetCell(COL_TYPE);
                if (typeCell != null)
                    typeCell.SetCellValue("–");

                ICell accCell = curRow.GetCell(COL_ACCEPT);
                if (accCell != null)
                    accCell.SetCellValue("ü");
            }
            catch (Exception ex)
            {
                logger.Debug(ex);
                throw;
            }
        }

        /// <summary>
        /// Merge the SN field of these specified rows
        /// </summary>
        /// <param name="rowStart"></param>
        /// <param name="rowEnd"></param>
        /// <exception cref="Exception"></exception>
        internal void MergeLineSN(int rowStart, int rowEnd)
        {
            if (rowStart >= PAGE_SIZE || rowEnd >= PAGE_SIZE)
                throw new Exception("MergeLineSN: rowStart and rowEnd must be less than PAGE_SIZE");

            if (rowStart > rowEnd)
                throw new Exception("MergeLineSN: rowStart must be less than rowEnd");

            // first is to make a check to make sure all the line has same SN
            IList<string> rowSNs = new List<string>();
            for (int i = rowStart; i <= rowEnd; i++)
            {
                IRow curRow = this.sheet.GetRow(OutReport.START_BODY_ROW + i);
                string sn = curRow.GetCell(COL_PART_SN).StringCellValue;
                rowSNs.Add(sn);
            }

            if (rowSNs.Distinct().Count() > 1)
                throw new Exception("MergeLineSN: rows must have same SN");

            var mergedRegions = this.sheet.MergedRegions;
            var regionsToRemove = new List<int>();

            // remove the merged range from A-L (COL_PART_SN to COL_EXTEND_OF_TEST - 1) in every specified row
            for (int i = rowStart; i <= rowEnd; i++)
            {
                // get the mergedRegion
                int mergedRegionIndex = mergedRegions.FindIndex(
                    x => x.FirstRow == OutReport.START_BODY_ROW + i &&
                    x.LastRow == OutReport.START_BODY_ROW + i &&
                    x.FirstColumn == COL_PART_SN &&
                    x.LastColumn == COL_EXTEND_OF_TEST - 1);
                regionsToRemove.Add(mergedRegionIndex);
            }

            if (regionsToRemove.Count > 0)
                this.sheet.RemoveMergedRegions(regionsToRemove);

            CellRangeAddress range = new CellRangeAddress(
                rowStart + START_BODY_ROW,
                rowEnd + START_BODY_ROW,
                COL_PART_SN,
                COL_EXTEND_OF_TEST - 1);
            this.sheet.AddMergedRegion(range);
        }
    }
}