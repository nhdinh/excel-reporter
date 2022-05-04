using System.Drawing;

namespace ExcelReporter
{
    public class AppConfig
    {
        public string LastReportOptionId { get; set; }

        public Point LastWindowLocation { get; set; }

        public Size LastWindowSize { get; set; }

        public string LastReportNo { get; internal set; }

        public HeaderField[] DefaultHeaderFields { get; set; }
    }
}