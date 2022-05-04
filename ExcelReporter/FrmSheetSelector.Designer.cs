namespace ExcelReporter
{
    partial class FrmSheetSelector
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.listSheetName = new System.Windows.Forms.ListView();
			this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// listSheetName
			// 
			this.listSheetName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.listSheetName.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
			this.listSheetName.FullRowSelect = true;
			this.listSheetName.GridLines = true;
			this.listSheetName.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
			this.listSheetName.HideSelection = false;
			this.listSheetName.Location = new System.Drawing.Point(12, 51);
			this.listSheetName.Name = "listSheetName";
			this.listSheetName.Size = new System.Drawing.Size(609, 344);
			this.listSheetName.TabIndex = 0;
			this.listSheetName.UseCompatibleStateImageBehavior = false;
			this.listSheetName.View = System.Windows.Forms.View.Details;
			this.listSheetName.ItemActivate += new System.EventHandler(this.listSheetName_ItemActivate);
			// 
			// columnHeader1
			// 
			this.columnHeader1.Width = 255;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(12, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 15);
			this.label1.TabIndex = 1;
			this.label1.Text = "label1";
			// 
			// FrmSheetSelector
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(633, 407);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.listSheetName);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "FrmSheetSelector";
			this.Text = "Select data sheet...";
			this.Load += new System.EventHandler(this.FrmSheetSelector_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView listSheetName;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.Label label1;
	}
}