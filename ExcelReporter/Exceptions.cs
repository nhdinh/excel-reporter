using System;
using System.Runtime.Serialization;

namespace ExcelReporter
{
    [Serializable]
    internal class SheetNotFoundException : Exception
    {
        private string path;
        private string sheetName;

        public SheetNotFoundException()
        {
        }

        public SheetNotFoundException(string message) : base(message)
        {
        }

        public SheetNotFoundException(string path, string sheetName)
        {
            this.path = path;
            this.sheetName = sheetName;
        }

        public SheetNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected SheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }

    [Serializable]
    internal class WrongReportFormatException : Exception
    {
        public WrongReportFormatException() : base("No DateTime column found. The opened file maybe in a wrong format.")
        {
        }
    }

    [Serializable]
    internal class HeaderRowNotFoundException : Exception
    {
        public HeaderRowNotFoundException() : base("Cannot detect Header Row of report file")
        {
        }
    }
}