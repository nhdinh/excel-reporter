namespace ExcelReporter
{
	partial class PendingLoadTabContent
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
			this.label1 = new System.Windows.Forms.Label();
			this.llblFilePath = new System.Windows.Forms.LinkLabel();
			this.lvSheetNames = new System.Windows.Forms.ListView();
			this.label2 = new System.Windows.Forms.Label();
			this.txtHeaderRowIndex = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.lbHeaderLabels = new System.Windows.Forms.Label();
			this.btnReload = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(13, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(64, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Loading file:";
			// 
			// llblFilePath
			// 
			this.llblFilePath.AutoSize = true;
			this.llblFilePath.Location = new System.Drawing.Point(123, 12);
			this.llblFilePath.Name = "llblFilePath";
			this.llblFilePath.Size = new System.Drawing.Size(55, 13);
			this.llblFilePath.TabIndex = 1;
			this.llblFilePath.TabStop = true;
			this.llblFilePath.Text = "linkLabel1";
			// 
			// lvSheetNames
			// 
			this.lvSheetNames.FullRowSelect = true;
			this.lvSheetNames.GridLines = true;
			this.lvSheetNames.HideSelection = false;
			this.lvSheetNames.Location = new System.Drawing.Point(126, 39);
			this.lvSheetNames.MultiSelect = false;
			this.lvSheetNames.Name = "lvSheetNames";
			this.lvSheetNames.Size = new System.Drawing.Size(290, 72);
			this.lvSheetNames.TabIndex = 2;
			this.lvSheetNames.UseCompatibleStateImageBehavior = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(13, 39);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(81, 13);
			this.label2.TabIndex = 3;
			this.label2.Text = "Selected sheet:";
			// 
			// txtHeaderRowIndex
			// 
			this.txtHeaderRowIndex.Location = new System.Drawing.Point(126, 117);
			this.txtHeaderRowIndex.Name = "txtHeaderRowIndex";
			this.txtHeaderRowIndex.Size = new System.Drawing.Size(290, 20);
			this.txtHeaderRowIndex.TabIndex = 4;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(13, 120);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(107, 13);
			this.label3.TabIndex = 5;
			this.label3.Text = "Header Row Number";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(13, 149);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(79, 13);
			this.label4.TabIndex = 6;
			this.label4.Text = "Header Labels:";
			// 
			// lbHeaderLabels
			// 
			this.lbHeaderLabels.Location = new System.Drawing.Point(123, 149);
			this.lbHeaderLabels.Name = "lbHeaderLabels";
			this.lbHeaderLabels.Size = new System.Drawing.Size(293, 69);
			this.lbHeaderLabels.TabIndex = 7;
			this.lbHeaderLabels.Text = "Header Labels";
			// 
			// btnReload
			// 
			this.btnReload.Location = new System.Drawing.Point(126, 221);
			this.btnReload.Name = "btnReload";
			this.btnReload.Size = new System.Drawing.Size(158, 23);
			this.btnReload.TabIndex = 8;
			this.btnReload.Text = "Correct and Reload";
			this.btnReload.UseVisualStyleBackColor = true;
			this.btnReload.Click += new System.EventHandler(this.btnReload_Click);
			// 
			// PendingLoadTabContent
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.btnReload);
			this.Controls.Add(this.lbHeaderLabels);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtHeaderRowIndex);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lvSheetNames);
			this.Controls.Add(this.llblFilePath);
			this.Controls.Add(this.label1);
			this.Name = "PendingLoadTabContent";
			this.Size = new System.Drawing.Size(704, 475);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.LinkLabel llblFilePath;
		private System.Windows.Forms.ListView lvSheetNames;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtHeaderRowIndex;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lbHeaderLabels;
		private System.Windows.Forms.Button btnReload;
	}
}
