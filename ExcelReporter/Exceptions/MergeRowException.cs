using System;
using System.Diagnostics;
using System.Runtime.Serialization;

namespace ExcelReporter.Exceptions
{
    [Serializable]
    [DebuggerDisplay("{" + nameof(GetDebuggerDisplay) + "(),nq}")]
    public class MergeRowException : ExcelReporterException
    {
        public MergeRowException()
        { }

        public MergeRowException(string message) : base(message)
        {
        }

        public MergeRowException(SerializationInfo info, StreamingContext context) : base(info, context)
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