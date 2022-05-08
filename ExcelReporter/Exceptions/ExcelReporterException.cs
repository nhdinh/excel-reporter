using System;
using System.Diagnostics;
using System.Runtime.Serialization;

namespace ExcelReporter.Exceptions
{
    [Serializable]
    [DebuggerDisplay("{" + nameof(GetDebuggerDisplay) + "(),nq}")]
    public class ExcelReporterException : Exception
    {
        public ExcelReporterException()
        {
        }

        public ExcelReporterException(string message) : base(message)
        {
        }

        protected ExcelReporterException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context) => base.GetObjectData(info, context);

        private string GetDebuggerDisplay()
        {
            return this.ToString();
        }
    }
}