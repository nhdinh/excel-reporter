using System;

namespace ExcelReporter.Exceptions
{
    [Serializable]
    public class WrongReportFormatException : Exception
    {
        public WrongReportFormatException() : base("No DateTime column found. The opened file maybe in a wrong format.")
        {
        }
    }

    [Serializable]
    public class HeaderRowNotFoundException : Exception
    {
        public HeaderRowNotFoundException() : base("Cannot detect Header Row of report file")
        {
        }
    }
}