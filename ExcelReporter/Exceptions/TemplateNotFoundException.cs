using System;
using System.Diagnostics;
using System.Runtime.Serialization;

namespace ExcelReporter.Exceptions
{
    [Serializable]
    [DebuggerDisplay("{" + nameof(GetDebuggerDisplay) + "(),nq}")]
    public class TemplateNotFoundException : ExcelReporterException
    {
        private readonly string templateForm;
        private readonly string path;

        public TemplateNotFoundException()
        {
        }

        public TemplateNotFoundException(string message) : base(message)
        {
        }

        public TemplateNotFoundException(string path, string templateForm)
        {
            this.path = path;
            this.templateForm = templateForm;
        }

        public TemplateNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context)
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