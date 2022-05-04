using NLog;
using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReporter
{
    public partial class FrmMain : Form
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private CoreApp app;

        public FrmMain()
        {
            InitializeComponent();
            Helpers.ConfigNLog();

            // starting by load working data from saved json
            this.app = CoreApp.GetInstance();

            // assign the handler for WorkingFilesChanged
            this.app.PropertyChanged += AppData_PropertyChanged;
        }

        /// <summary>
        /// Handle PropertyChanged event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AppData_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            Helpers.SaveRecentFile(this.app.WorkFiles);

            // what report was changed
            var changedReportTag = e.PropertyName;

            // clear all tabs
            if (changedReportTag != "")
                foreach (TabPage tab in this.tabs.TabPages)
                {
                    if ((string)tab.Tag == changedReportTag)
                    {
                        this.tabs.TabPages.Remove(tab);
                    }
                }
            else
                this.tabs.TabPages.Clear();

            // make tab and add to TabsContainer
            if (changedReportTag == "")
                Parallel.ForEach(app.Reports, report =>
                {
                    if (report.LoadStatus == LoadStatus.LOADED)
                    {
                        TabPage page = this.makeTabPage(report);
                        this.tabs.TabPages.Add(page);
                    }
                    else
                    {
                        FrmSheetSelector frmSheetSelector = new FrmSheetSelector(report.WorkFile.GetFileName(), report.SheetNames);
                        DialogResult dlgRes = frmSheetSelector.ShowDialog();
                        if (dlgRes == DialogResult.OK)
                        {
                            string selectedSheetName = frmSheetSelector.SelectedSheet;

                            // update the selected SheetName into workfile
                            var workFile = report.WorkFile;
                            workFile.SheetName = selectedSheetName;

                            this.app.UpdateWorkfile(workFile);
                        }
                    }
                });
            else
            {
                try
                {
                    var report = this.app.Reports.Where(r => r.Tag == changedReportTag).First();
                    TabPage page = this.makeTabPage(report);
                    this.tabs.TabPages.Add(page);
                }
                finally { }
            }
        }

        /// <summary>
        /// handle actions when form is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                this.app.Config.LastWindowSize = this.Size;
                this.app.Config.LastWindowLocation = this.Location;

                Helpers.SaveRecentFile(this.app.WorkFiles);
                Helpers.SaveAppConfig(this.app.Config);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            // set window size and location
            if (this.app.Config.LastWindowLocation != null)
            {
                if (this.app.Config.LastWindowLocation.X < Screen.PrimaryScreen.Bounds.Width && this.app.Config.LastWindowLocation.Y < Screen.PrimaryScreen.Bounds.Height)
                    this.Location = this.app.Config.LastWindowLocation;
            }
            else
            {
                this.StartPosition = FormStartPosition.CenterScreen;
            }

            if (this.app.Config.LastWindowSize != null)
            {
                this.Size = this.app.Config.LastWindowSize;
            }
            else
            {
                this.Size = new Size(697, 403);
            }
        }

        /// <summary>
        /// Generate TabPage with DataReport
        /// </summary>
        /// <param name="report"></param>
        /// <returns></returns>
        private TabPage makeTabPage(DataReport report)
        {
            TabContent tabContent = new TabContent(report);
            var reportInfo = report.GenerateWorkFile();

            TabPage tabPage = new TabPage(String.Format("{0} - {1}", reportInfo.GetFileName(), reportInfo.SheetName));
            tabPage.Tag = report.Tag;
            tabPage.Controls.Add(tabContent);
            //tabContent.Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Left;
            tabContent.Dock = DockStyle.Fill;

            return tabPage;
        }

        private void openFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Office Excel Files|*.xls;*.xlsx"
            };
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                var fileNames = openFileDialog.FileNames;
                var _workingFiles = new WorkFileInfo[fileNames.Length];

                Parallel.For(0, fileNames.Length, fileIndex =>
                {
                    _workingFiles[fileIndex] = new WorkFileInfo()
                    {
                        FilePath = fileNames[fileIndex],
                        SheetName = "",
                        SheetNames = null,
                        HeaderRowIndex = -1,
                        HeaderLabels = null
                    };
                });

                Logger.Debug("Set {0} working files", _workingFiles.Length);
                this.app.LoadWorkFiles(_workingFiles);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void exportFromOpenFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmExportConfig frm = new FrmExportConfig();
            DialogResult res = frm.ShowDialog();
            if (res == DialogResult.OK)
            {
            }
        }
    }
}