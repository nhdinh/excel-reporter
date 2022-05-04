using System.Linq;
using System.Windows.Forms;

namespace ExcelReporter
{
    public partial class TabContent : UserControl
    {
        private readonly DataReport _report;

        public TabContent(DataReport report)
        {
            InitializeComponent();
            this._report = report;
            this.Tag = this._report.Id;

            this.pagination.PropertyChanged += Pagination_PropertyChanged;
            this.pagination.PageSize = 50;
            this.pagination.TotalItems = this._report.TotalItems;
        }

        private void Pagination_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            int firstIndex = (pagination.CurrentPage - 1) * pagination.PageSize;
            int lastIndex = pagination.CurrentPage * pagination.PageSize;

            if (lastIndex > this._report.ReportData.Count)
                lastIndex = this._report.ReportData.Count;

            // disable the header
            var data = this._report.ReportData.Skip(firstIndex).Take(lastIndex - firstIndex).ToList();
            this.dataViewer.AutoGenerateColumns = true;
            this.dataViewer.DataSource = data;
        }
    }
}