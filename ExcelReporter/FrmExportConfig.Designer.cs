using System;
using System.ComponentModel;

namespace ExcelReporter
{
	partial class FrmExportConfig
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.dgParts = new System.Windows.Forms.DataGridView();
            this.btnDelete = new System.Windows.Forms.Button();
            this.cbSelectReportOptionId = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtAcceptanceCriteria = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtProcedure = new System.Windows.Forms.TextBox();
            this.txtCustomerName = new System.Windows.Forms.TextBox();
            this.btnSaveAndClose = new System.Windows.Forms.Button();
            this.btnSaveAndExport = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.mt202 = new System.Windows.Forms.RadioButton();
            this.mt201 = new System.Windows.Forms.RadioButton();
            this.label4 = new System.Windows.Forms.Label();
            this.dtpEnd = new System.Windows.Forms.DateTimePicker();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.dtpStart = new System.Windows.Forms.DateTimePicker();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgParts)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Customer name:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Procedure";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.dgParts);
            this.groupBox1.Controls.Add(this.btnDelete);
            this.groupBox1.Controls.Add(this.cbSelectReportOptionId);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.txtAcceptanceCriteria);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtProcedure);
            this.groupBox1.Controls.Add(this.txtCustomerName);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(587, 577);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Report Settings";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(4, 137);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(102, 13);
            this.label11.TabIndex = 22;
            this.label11.Text = "Parts description";
            // 
            // dgParts
            // 
            this.dgParts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgParts.Location = new System.Drawing.Point(4, 161);
            this.dgParts.Name = "dgParts";
            this.dgParts.RowHeadersWidth = 51;
            this.dgParts.Size = new System.Drawing.Size(575, 410);
            this.dgParts.TabIndex = 21;
            this.dgParts.VirtualMode = true;
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(473, 24);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(108, 23);
            this.btnDelete.TabIndex = 20;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseMnemonic = false;
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // cbSelectReportOptionId
            // 
            this.cbSelectReportOptionId.FormattingEnabled = true;
            this.cbSelectReportOptionId.Items.AddRange(new object[] {
            "New..."});
            this.cbSelectReportOptionId.Location = new System.Drawing.Point(137, 26);
            this.cbSelectReportOptionId.Name = "cbSelectReportOptionId";
            this.cbSelectReportOptionId.Size = new System.Drawing.Size(330, 21);
            this.cbSelectReportOptionId.TabIndex = 0;
            this.cbSelectReportOptionId.SelectedIndexChanged += new System.EventHandler(this.cbSelectReportOptionId_SelectedIndexChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 29);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(57, 13);
            this.label10.TabIndex = 10;
            this.label10.Text = "Setting ID:";
            // 
            // txtAcceptanceCriteria
            // 
            this.txtAcceptanceCriteria.Location = new System.Drawing.Point(137, 105);
            this.txtAcceptanceCriteria.Name = "txtAcceptanceCriteria";
            this.txtAcceptanceCriteria.Size = new System.Drawing.Size(444, 20);
            this.txtAcceptanceCriteria.TabIndex = 3;
            this.txtAcceptanceCriteria.Validating += new System.ComponentModel.CancelEventHandler(this.txtAcceptanceCriteria_Validating);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 108);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Acceptance Criteria";
            // 
            // txtProcedure
            // 
            this.txtProcedure.Location = new System.Drawing.Point(137, 79);
            this.txtProcedure.Name = "txtProcedure";
            this.txtProcedure.Size = new System.Drawing.Size(444, 20);
            this.txtProcedure.TabIndex = 2;
            this.txtProcedure.Validating += new System.ComponentModel.CancelEventHandler(this.txtProcedure_Validating);
            // 
            // txtCustomerName
            // 
            this.txtCustomerName.Location = new System.Drawing.Point(137, 53);
            this.txtCustomerName.Name = "txtCustomerName";
            this.txtCustomerName.Size = new System.Drawing.Size(444, 20);
            this.txtCustomerName.TabIndex = 1;
            this.txtCustomerName.Validating += new System.ComponentModel.CancelEventHandler(this.txtCustomerName_Validating);
            // 
            // btnSaveAndClose
            // 
            this.btnSaveAndClose.Location = new System.Drawing.Point(363, 691);
            this.btnSaveAndClose.Name = "btnSaveAndClose";
            this.btnSaveAndClose.Size = new System.Drawing.Size(115, 36);
            this.btnSaveAndClose.TabIndex = 3;
            this.btnSaveAndClose.Text = "Save & Close";
            this.btnSaveAndClose.UseMnemonic = false;
            this.btnSaveAndClose.UseVisualStyleBackColor = true;
            this.btnSaveAndClose.Click += new System.EventHandler(this.btnSaveAndClose_Click);
            // 
            // btnSaveAndExport
            // 
            this.btnSaveAndExport.Location = new System.Drawing.Point(166, 691);
            this.btnSaveAndExport.Name = "btnSaveAndExport";
            this.btnSaveAndExport.Size = new System.Drawing.Size(191, 36);
            this.btnSaveAndExport.TabIndex = 2;
            this.btnSaveAndExport.Text = "Save & Export report";
            this.btnSaveAndExport.UseMnemonic = false;
            this.btnSaveAndExport.UseVisualStyleBackColor = true;
            this.btnSaveAndExport.Click += new System.EventHandler(this.btnSaveAndExport_Click);
            // 
            // btnClose
            // 
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(484, 691);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(115, 36);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.UseMnemonic = false;
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.mt202);
            this.groupBox2.Controls.Add(this.mt201);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.dtpEnd);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.dtpStart);
            this.groupBox2.Location = new System.Drawing.Point(12, 595);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(587, 90);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Report Date, Number";
            // 
            // mt202
            // 
            this.mt202.AutoSize = true;
            this.mt202.Location = new System.Drawing.Point(182, 57);
            this.mt202.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.mt202.Name = "mt202";
            this.mt202.Size = new System.Drawing.Size(59, 17);
            this.mt202.TabIndex = 15;
            this.mt202.Text = "MT202";
            this.mt202.UseVisualStyleBackColor = true;
            this.mt202.CheckedChanged += new System.EventHandler(this.mt202_CheckedChanged);
            // 
            // mt201
            // 
            this.mt201.AutoSize = true;
            this.mt201.Checked = true;
            this.mt201.Location = new System.Drawing.Point(81, 58);
            this.mt201.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.mt201.Name = "mt201";
            this.mt201.Size = new System.Drawing.Size(59, 17);
            this.mt201.TabIndex = 14;
            this.mt201.TabStop = true;
            this.mt201.Text = "MT201";
            this.mt201.UseVisualStyleBackColor = true;
            this.mt201.CheckedChanged += new System.EventHandler(this.mt201_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 60);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Report Type";
            // 
            // dtpTo
            // 
            this.dtpEnd.CustomFormat = "dd/ MMM/ yyyy";
            this.dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpEnd.Location = new System.Drawing.Point(381, 31);
            this.dtpEnd.Name = "dtpTo";
            this.dtpEnd.Size = new System.Drawing.Size(200, 20);
            this.dtpEnd.TabIndex = 13;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(299, 37);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(23, 13);
            this.label13.TabIndex = 12;
            this.label13.Text = "To:";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(6, 32);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(33, 13);
            this.label12.TabIndex = 11;
            this.label12.Text = "From:";
            // 
            // dtpFrom
            // 
            this.dtpStart.CustomFormat = "dd/ MMM/ yyyy";
            this.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpStart.Location = new System.Drawing.Point(81, 31);
            this.dtpStart.Name = "dtpFrom";
            this.dtpStart.Size = new System.Drawing.Size(200, 20);
            this.dtpStart.TabIndex = 0;
            this.dtpStart.ValueChanged += new System.EventHandler(this.dtpFrom_ValueChanged);
            // 
            // FrmExportConfig
            // 
            this.AcceptButton = this.btnSaveAndExport;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(611, 739);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSaveAndExport);
            this.Controls.Add(this.btnSaveAndClose);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmExportConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmExportConfig";
            this.Load += new System.EventHandler(this.frmExportConfig_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgParts)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.TextBox txtAcceptanceCriteria;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtProcedure;
		private System.Windows.Forms.TextBox txtCustomerName;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.ComboBox cbSelectReportOptionId;
		private System.Windows.Forms.Button btnSaveAndClose;
		private System.Windows.Forms.Button btnSaveAndExport;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnDelete;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.DataGridView dgParts;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.DateTimePicker dtpStart;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.DateTimePicker dtpEnd;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RadioButton mt202;
        private System.Windows.Forms.RadioButton mt201;
    }
}