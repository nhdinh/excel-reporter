using NPOI.SS.UserModel;
using System;

namespace ExcelReporter
{
    public struct CellAddress
    {
        public int Row;
        public int Col;

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
        public DateTime InspectedDate;
        public string PartNo;
        public double Concentration;
        public string Inspector;
        public string Shift;

        public ReportKey(DateTime inspectedDate, string partNo, double concentration, string inspector, string shift)
        {
            this.InspectedDate = inspectedDate;
            this.PartNo = partNo;
            this.Concentration = concentration;
            this.Inspector = inspector;
            this.Shift = shift;
        }
    }

    public static class CELL_ADDRESSES
    {
        public static CellAddress REPORT_NO = new CellAddress(3, "AL");

        public static CellAddress CUSTOMER_NAME = new CellAddress(5, "A");
        public static CellAddress PART_NAME = new CellAddress(5, "V");
        public static CellAddress PART_NO = new CellAddress(5, "AL");
        public static CellAddress PROCEDURE = new CellAddress(7, "V");
        public static CellAddress ACCEPTANCE_CRITERIA = new CellAddress(7, "AL");
        public static CellAddress TEST_TEMP = new CellAddress(9, "V");
        public static CellAddress CONCENTRATION = new CellAddress(17, "H");
        public static CellAddress COIL = new CellAddress(15, "AO");
        public static CellAddress YOKE = new CellAddress(16, "AO");

        public static CellAddress BLACK_LIGHT_MODEL = new CellAddress(15, "AA");
        public static CellAddress BLACK_LIGHT_SN = new CellAddress(16, "X");
        public static CellAddress UV = new CellAddress(17, "Z");

        public static CellAddress VISIBLE_LIGHT_INTENSITY = new CellAddress(19, "F");

        public static CellAddress DATE_OF_EXAMINATION = new CellAddress(45, "AO");
        public static CellAddress DATE_OF_REPORT = new CellAddress(46, "AO");
        public static CellAddress PAGE_OF_TOTAL = new CellAddress(47, "AY");

        public static CellAddress INSPECTOR = new CellAddress(46, "K");
    }

    public class ExcelReport
    {
        public ISheet Data { get; set; }
        public static int START_BODY_ROW = 24;
        public static int PAGE_SIZE = 18;

        public void SetFieldValue(CellAddress targetCell, object value)
        {
            if (this.Data == null)
            {
                throw new Exception("Must set Sheet instance first");
            }

            ICell cell = this.Data.GetRow(targetCell.Row).GetCell(targetCell.Col);

            if (value.GetType() == typeof(string))
            {
                cell.SetCellValue((string)value);
                cell.SetCellType(CellType.String);
                return;
            }
            else if (value.GetType() == typeof(DateTime))
            {
                cell.SetCellValue((DateTime)value);
                cell.SetCellType(CellType.Numeric);
                return;
            }
            else if (value.GetType() == typeof(int))
            {
                cell.SetCellValue((int)value);
                cell.SetCellType(CellType.Numeric);
                return;
            }
            else if (value.GetType() == typeof(double))
            {
                cell.SetCellValue((double)value);
                cell.SetCellType(CellType.Numeric);
                return;
            }
        }

        public void SetCustomerName(string customerName)
        {
            this.SetFieldValue(CELL_ADDRESSES.CUSTOMER_NAME, customerName);
        }

        public void SetProcedure(string procedure)
        {
            this.SetFieldValue(CELL_ADDRESSES.PROCEDURE, procedure);
        }

        internal void SetLineValue(int row, string partSN, string extendOfTest)
        {
            if (row >= PAGE_SIZE)
                row -= PAGE_SIZE;

            try
            {
                IRow curRow = this.Data.GetRow(ExcelReport.START_BODY_ROW + row);

                ICell snCell = curRow.GetCell(0);
                if (snCell != null)
                    snCell.SetCellValue(partSN);

                ICell extendOfTestCell = curRow.GetCell(12);
                if (extendOfTestCell != null)
                    extendOfTestCell.SetCellValue(extendOfTest);

                ICell qtyCell = curRow.GetCell(18);
                if (qtyCell != null)
                    qtyCell.SetCellValue(1);

                ICell locationCell = curRow.GetCell(21);
                if (locationCell != null)
                    locationCell.SetCellValue("–");

                ICell lengthCell = curRow.GetCell(24);
                if (lengthCell != null)
                    lengthCell.SetCellValue("–");

                ICell widthCell = curRow.GetCell(27);
                if (widthCell != null)
                    widthCell.SetCellValue("–");

                ICell depthCell = curRow.GetCell(30);
                if (depthCell != null)
                    depthCell.SetCellValue("–");

                ICell areaCell = curRow.GetCell(33);
                if (areaCell != null)
                    areaCell.SetCellValue("–");

                ICell typeCell = curRow.GetCell(37);
                if (typeCell != null)
                    typeCell.SetCellValue("–");

                ICell accCell = curRow.GetCell(40);
                if (accCell != null)
                    accCell.SetCellValue("ü");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}