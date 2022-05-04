using System.Collections.Generic;

namespace ExcelReporter
{
    public class ReportOption
    {
        public string Id { get; set; } = "";
        public string CustomerName { get; set; } = "";
        public string Procedure { get; set; } = "";
        public string AcceptanceCriteria { get; set; } = "";
        public IList<Part> Parts { get; set; } = new List<Part>();
    }

    public class Part
    {
        public string PartNo { get; set; }
        public string Description { get; set; }
        public int Coil { get; set; }
        public int Yoke { get; set; }
        public string ExtendOfTest { get; set; }
    }
}