using System;
using System.Windows.Forms;

namespace ExcelReporter
{
	public partial class PendingLoadTabContent : UserControl
	{
		private WorkFileInfo workFile;

		public WorkFileInfo GetWorkFile
		{
			get
			{
				return workFile;
			}
		}


		public event System.EventHandler WorkFileChanged;

		public virtual void OnWorkFileChanged(WorkFileInfo workFile)
		{
			if (WorkFileChanged != null)
				WorkFileChanged(workFile, EventArgs.Empty);
		}

		public PendingLoadTabContent(WorkFileInfo workFile)
		{
			InitializeComponent();
			this.WorkFileChanged += PendingLoadTabContent_WorkFileChanged;

			this.workFile = workFile;
			this.llblFilePath.Text = workFile.FilePath;
			this.txtHeaderRowIndex.Text = workFile.HeaderRowIndex.ToString();

			var sheetNames = DataReport.GetSheetNames(workFile.FilePath);
			this.lvSheetNames.Items.Clear();
			foreach (var sheetName in sheetNames)
				this.lvSheetNames.Items.Add(sheetName);

			this.lbHeaderLabels.Text = this.buildHeaderLabels(workFile.HeaderLabels);

		}

		private void PendingLoadTabContent_WorkFileChanged(object sender, EventArgs e)
		{
			var newWorkFile = (WorkFileInfo)sender;

			if (newWorkFile != null && newWorkFile != this.workFile)
			{
				this.workFile = newWorkFile;
				this.llblFilePath.Text = newWorkFile.FilePath;
				this.lvSheetNames.Items.Clear();
				this.txtHeaderRowIndex.Text = newWorkFile.HeaderRowIndex.ToString();
				this.lbHeaderLabels.Text = this.buildHeaderLabels(newWorkFile.HeaderLabels);
			}
		}

		private string buildHeaderLabels(string[] labels)
		{
			if (labels == null)
				return string.Empty;

			return String.Join("; ", labels);
		}

		private void btnReload_Click(object sender, EventArgs e)
		{
			try
			{
				var sheetName = this.lvSheetNames.SelectedItems[0].Text;
				var headerRowIndex = -1;
				int.TryParse(this.txtHeaderRowIndex.Text, out headerRowIndex);

				var report = DataReport.LoadFile(this.workFile.FilePath, sheetName);
				var newWorkFile = new WorkFileInfo();
				newWorkFile.FilePath = report.FilePath;
				newWorkFile.SheetName = sheetName;
				newWorkFile.HeaderRowIndex = report.HeaderRowIndex;
				newWorkFile.HeaderLabels = report.HeaderDatas;

				OnWorkFileChanged(newWorkFile);
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.ToString());
			}
		}
	}
}
