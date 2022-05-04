namespace ExcelReporter
{
    partial class TabContent
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataViewer = new System.Windows.Forms.DataGridView();
            this.pagination = new ExcelReporter.HorizontalPagination();
            ((System.ComponentModel.ISupportInitialize)(this.dataViewer)).BeginInit();
            this.SuspendLayout();
            // 
            // dataViewer
            // 
            this.dataViewer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataViewer.ColumnHeadersHeight = 29;
            this.dataViewer.Location = new System.Drawing.Point(4, 4);
            this.dataViewer.Margin = new System.Windows.Forms.Padding(4);
            this.dataViewer.Name = "dataViewer";
            this.dataViewer.ReadOnly = true;
            this.dataViewer.RowHeadersWidth = 72;
            this.dataViewer.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataViewer.ShowCellErrors = false;
            this.dataViewer.ShowCellToolTips = false;
            this.dataViewer.ShowEditingIcon = false;
            this.dataViewer.ShowRowErrors = false;
            this.dataViewer.Size = new System.Drawing.Size(1193, 682);
            this.dataViewer.TabIndex = 2;
            // 
            // pagination
            // 
            this.pagination.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pagination.CurrentPage = 0;
            this.pagination.Location = new System.Drawing.Point(4, 693);
            this.pagination.Name = "pagination";
            this.pagination.PageSize = 10;
            this.pagination.Size = new System.Drawing.Size(1193, 36);
            this.pagination.TabIndex = 3;
            this.pagination.TotalItems = 0;
            // 
            // TabContent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pagination);
            this.Controls.Add(this.dataViewer);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TabContent";
            this.Size = new System.Drawing.Size(1201, 742);
            ((System.ComponentModel.ISupportInitialize)(this.dataViewer)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataViewer;
        private HorizontalPagination pagination;
    }
}
