using System;
using System.Linq;
using System.Windows.Forms;

namespace ExcelReporter
{
    public partial class FrmSheetSelector : Form
    {
        private string[] sheetNames = null;
        private string selectedSheet;
        private string fileName;

        public FrmSheetSelector(string fileName, string[] SelectableSheetNames)
        {
            InitializeComponent();
            this.sheetNames = SelectableSheetNames;
            this.fileName = fileName;
        }

        public void setSheetName(string[] sheetNames)
        {
            this.sheetNames = sheetNames;
        }

        public string SelectedSheet
        {
            get
            {
                return this.selectedSheet;
            }
        }

        private void FrmSheetSelector_Load(object sender, EventArgs e)
        {
            this.label1.Text = string.Format("Select the working sheet of file {0}", this.fileName);

            if (this.sheetNames == null)
                this.sheetNames = new string[0];

            this.listSheetName.Items.AddRange(this.sheetNames.Select(name => new ListViewItem(name)).ToArray());
        }

        private void listSheetName_ItemActivate(object sender, EventArgs e)
        {
            this.selectedSheet = this.sheetNames[
                this.listSheetName.SelectedIndices[0]
                ];
            this.DialogResult = DialogResult.OK;
        }
    }
}