using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReporter.Exceptions;
using NLog;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

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
        public int Row { get; }
        public int Col { get; }

        public CellAddress(int row, string col)
        {
            Row = row - 1;
            Col = GetColumnNumber(col) - 1;
        }

        private static int GetColumnNumber(string name)
        {
            var number = 0;
            var pow = 1;
            for (var i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }
    }

    public struct ReportKey
    {
        public DateTime InspectedDate { get; }
        public string PartNo { get; }
        public double Concentration { get; }
        public string Inspector { get; }
        public string Shift { get; }

        public ReportKey(DateTime inspectedDate, string partNo, double concentration, string inspector, string shift)
        {
            InspectedDate = inspectedDate;
            PartNo = partNo;
            Concentration = concentration;
            Inspector = inspector;
            Shift = shift;
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

        public const int START_BODY_ROW = 24;
        public const int PAGE_SIZE = 18;
        private readonly Logger logger;

        public OutReport(ISheet newSheet)
        {
            sheet = newSheet;
            logger = LogManager.GetCurrentClassLogger();
        }

        private ISheet sheet { get; }

        public void SetFieldValue(CellAddress targetCell, object value)
        {
            if (sheet == null) throw new Exception("Must set Sheet instance first");

            var cell = sheet.GetRow(targetCell.Row).GetCell(targetCell.Col);

            switch (value.GetType())
            {
                case Type t when t == typeof(string):
                    cell.SetCellValue(value.ToString());
                    cell.SetCellType(CellType.String);
                    break;
                case Type t when t == typeof(DateTime):
                    cell.SetCellValue(Convert.ToDateTime(value));
                    cell.SetCellType(CellType.Numeric);
                    break;
                case Type t when t == typeof(int):
                    cell.SetCellValue(Convert.ToInt32(value));
                    cell.SetCellType(CellType.Numeric);
                    break;
                case Type t when t == typeof(double):
                    cell.SetCellValue(Convert.ToDouble(value));
                    cell.SetCellType(CellType.Numeric);
                    break;
                case Type t when t == typeof(float) || t == typeof(System.Single):
                    cell.SetCellValue(Convert.ToDouble(value));
                    cell.SetCellType(CellType.Numeric);
                    break;
                default:
                    cell.SetCellValue(value.ToString());
                    cell.SetCellType(CellType.String);
                    break;
            }
        }

        public void SetCustomerName(string customerName)
        {
            SetFieldValue(CellLabels.CUSTOMER_NAME, customerName);
        }

        public void SetProcedure(string procedure)
        {
            SetFieldValue(CellLabels.PROCEDURE, procedure);
        }

        internal void SetLineValue(int row, string partSN, string extendOfTest)
        {
            if (row >= PAGE_SIZE)
                row -= PAGE_SIZE;

            try
            {
                var curRow = sheet.GetRow(START_BODY_ROW + row);

                var snCell = curRow.GetCell(COL_PART_SN);
                if (snCell != null)
                    snCell.SetCellValue(partSN);

                var extendOfTestCell = curRow.GetCell(COL_EXTEND_OF_TEST);
                if (extendOfTestCell != null)
                    extendOfTestCell.SetCellValue(extendOfTest);

                var qtyCell = curRow.GetCell(COL_QUANTITY);
                if (qtyCell != null)
                    qtyCell.SetCellValue(1);

                var locationCell = curRow.GetCell(COL_LOCATION);
                if (locationCell != null)
                    locationCell.SetCellValue("–");

                var lengthCell = curRow.GetCell(COL_LENGTH);
                if (lengthCell != null)
                    lengthCell.SetCellValue("–");

                var widthCell = curRow.GetCell(COL_WIDTH);
                if (widthCell != null)
                    widthCell.SetCellValue("–");

                var depthCell = curRow.GetCell(COL_DEPTH);
                if (depthCell != null)
                    depthCell.SetCellValue("–");

                var areaCell = curRow.GetCell(COL_AREA);
                if (areaCell != null)
                    areaCell.SetCellValue("–");

                var typeCell = curRow.GetCell(COL_TYPE);
                if (typeCell != null)
                    typeCell.SetCellValue("–");

                var accCell = curRow.GetCell(COL_ACCEPT);
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
        ///     Merge the SN field of these specified rows
        /// </summary>
        /// <param name="rowStart"></param>
        /// <param name="rowEnd"></param>
        /// <exception cref="Exception"></exception>
        internal void MergeLineSN(int rowStart, int rowEnd)
        {
            if (rowStart >= PAGE_SIZE || rowEnd >= PAGE_SIZE)
                throw new MergeRowException("rowStart and rowEnd must be less than PAGE_SIZE");

            if (rowStart > rowEnd)
                throw new MergeRowException("rowStart must be less than rowEnd");

            // first is to make a check to make sure all the line has same SN
            IList<string> rowSNs = new List<string>();
            for (var i = rowStart; i <= rowEnd; i++)
            {
                var curRow = sheet.GetRow(START_BODY_ROW + i);
                var sn = curRow.GetCell(COL_PART_SN).StringCellValue;
                rowSNs.Add(sn);
            }

            if (rowSNs.Distinct().Count() > 1)
                throw new MergeRowException("Rows must have same SN");

            var mergedRegions = sheet.MergedRegions;
            var regionsToRemove = new List<int>();

            // remove the merged range from A-L (COL_PART_SN to COL_EXTEND_OF_TEST - 1) in every specified row
            for (var i = rowStart; i <= rowEnd; i++)
            {
                // get the mergedRegion
                var mergedRegionIndex = mergedRegions.FindIndex(
                    x => x.FirstRow == START_BODY_ROW + i &&
                         x.LastRow == START_BODY_ROW + i &&
                         x.FirstColumn == COL_PART_SN &&
                         x.LastColumn == COL_EXTEND_OF_TEST - 1);
                regionsToRemove.Add(mergedRegionIndex);
            }

            if (regionsToRemove.Count > 0)
                sheet.RemoveMergedRegions(regionsToRemove);

            var range = new CellRangeAddress(
                rowStart + START_BODY_ROW,
                rowEnd + START_BODY_ROW,
                COL_PART_SN,
                COL_EXTEND_OF_TEST - 1);
            sheet.AddMergedRegion(range);
        }
    }
}