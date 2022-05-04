using System;
using System.Windows.Forms;

namespace ExcelReporter
{
	public partial class TabLoader : UserControl
	{
		private WorkFileInfo currentWorkFile = null;
		private string loaderTitle = string.Empty;

		public WorkFileInfo WorkFile
		{
			get; set;
		}

		public string LoaderTitle
		{
			get;
		}

		public TabLoader(WorkFileInfo workFile)
		{
			InitializeComponent();
			this.currentWorkFile = workFile;
		}

		private void TabLoader_Load(object sender, System.EventArgs e)
		{
			UserControl tabContent = new PendingLoadTabContent(currentWorkFile);

			try
			{
				// make report and inject into TabContent
				DataReport report = DataReport.LoadFile(currentWorkFile.FilePath, currentWorkFile.SheetName);
				tabContent = new TabContent(report);

				this.currentWorkFile = report.GenerateWorkFile();
				this.loaderTitle = String.Format("{0} - {1}",
					currentWorkFile.GetFileName(), report.LoadStatus);
			}
			catch (Exception ex)
			{
				throw ex;
			}


		}
	}
}
