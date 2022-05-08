using System;
using System.Diagnostics;
using System.Runtime.Serialization;

namespace ExcelReporter.Exceptions
{
    [Serializable]
    [DebuggerDisplay("{" + nameof(GetDebuggerDisplay) + "(),nq}")]
    public class CreateNewSheetException : ExcelReporterException
    {
        private readonly string path;
        private readonly string sheetName;

        public CreateNewSheetException()
        {
        }

        public CreateNewSheetException(string message) : base(message)
        {
        }

        public CreateNewSheetException(string path, string sheetName)
        {
            this.path = path;
            this.sheetName = sheetName;
        }

        public CreateNewSheetException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }

        private string GetDebuggerDisplay()
        {
            return ToString();
        }
    }
}