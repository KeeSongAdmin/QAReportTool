
namespace QAReportTool
{
	partial class FormFreshWholeChicken
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
			this.btn_Day = new System.Windows.Forms.Button();
			this.btn_Month = new System.Windows.Forms.Button();
			this.comboBoxMonth = new System.Windows.Forms.ComboBox();
			this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
			this.comboBoxYear = new System.Windows.Forms.ComboBox();
			this.lb_TemplatePath = new System.Windows.Forms.Label();
			this.lb_OutputPath = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btn_Day
			// 
			this.btn_Day.Location = new System.Drawing.Point(322, 10);
			this.btn_Day.Name = "btn_Day";
			this.btn_Day.Size = new System.Drawing.Size(160, 29);
			this.btn_Day.TabIndex = 2;
			this.btn_Day.Text = "Generate Day Report";
			this.btn_Day.UseVisualStyleBackColor = true;
			this.btn_Day.Click += new System.EventHandler(this.btn_Day_Click);
			// 
			// btn_Month
			// 
			this.btn_Month.Location = new System.Drawing.Point(322, 43);
			this.btn_Month.Name = "btn_Month";
			this.btn_Month.Size = new System.Drawing.Size(160, 48);
			this.btn_Month.TabIndex = 2;
			this.btn_Month.Text = "Generate Month Report";
			this.btn_Month.UseVisualStyleBackColor = true;
			this.btn_Month.Click += new System.EventHandler(this.btn_Month_Click);
			// 
			// comboBoxMonth
			// 
			this.comboBoxMonth.FormattingEnabled = true;
			this.comboBoxMonth.Items.AddRange(new object[] {
            "Please Select Month",
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December"});
			this.comboBoxMonth.Location = new System.Drawing.Point(12, 70);
			this.comboBoxMonth.Name = "comboBoxMonth";
			this.comboBoxMonth.Size = new System.Drawing.Size(304, 21);
			this.comboBoxMonth.TabIndex = 3;
			this.comboBoxMonth.SelectedIndexChanged += new System.EventHandler(this.comboBoxMonth_SelectedIndexChanged);
			// 
			// dateTimePicker1
			// 
			this.dateTimePicker1.Location = new System.Drawing.Point(12, 12);
			this.dateTimePicker1.Name = "dateTimePicker1";
			this.dateTimePicker1.Size = new System.Drawing.Size(304, 20);
			this.dateTimePicker1.TabIndex = 4;
			// 
			// comboBoxYear
			// 
			this.comboBoxYear.FormattingEnabled = true;
			this.comboBoxYear.Items.AddRange(new object[] {
            "Please Select Month",
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December"});
			this.comboBoxYear.Location = new System.Drawing.Point(12, 43);
			this.comboBoxYear.Name = "comboBoxYear";
			this.comboBoxYear.Size = new System.Drawing.Size(304, 21);
			this.comboBoxYear.TabIndex = 3;
			this.comboBoxYear.SelectedIndexChanged += new System.EventHandler(this.comboBoxYear_SelectedIndexChanged);
			// 
			// lb_TemplatePath
			// 
			this.lb_TemplatePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lb_TemplatePath.ForeColor = System.Drawing.SystemColors.ActiveCaption;
			this.lb_TemplatePath.Location = new System.Drawing.Point(2, 94);
			this.lb_TemplatePath.Name = "lb_TemplatePath";
			this.lb_TemplatePath.Size = new System.Drawing.Size(480, 43);
			this.lb_TemplatePath.TabIndex = 5;
			this.lb_TemplatePath.Text = "Template Path : ";
			this.lb_TemplatePath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lb_OutputPath
			// 
			this.lb_OutputPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lb_OutputPath.ForeColor = System.Drawing.Color.DarkGreen;
			this.lb_OutputPath.Location = new System.Drawing.Point(2, 137);
			this.lb_OutputPath.Name = "lb_OutputPath";
			this.lb_OutputPath.Size = new System.Drawing.Size(480, 66);
			this.lb_OutputPath.TabIndex = 6;
			this.lb_OutputPath.Text = "Output Path : ";
			this.lb_OutputPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// FormFreshWholeChicken
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(484, 211);
			this.Controls.Add(this.lb_OutputPath);
			this.Controls.Add(this.lb_TemplatePath);
			this.Controls.Add(this.dateTimePicker1);
			this.Controls.Add(this.comboBoxYear);
			this.Controls.Add(this.comboBoxMonth);
			this.Controls.Add(this.btn_Month);
			this.Controls.Add(this.btn_Day);
			this.Name = "FormFreshWholeChicken";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "STOCK QC FORM 货品品质检查表_Fresh Whole Chicken";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}

		#endregion
		private System.Windows.Forms.Button btn_Day;
		private System.Windows.Forms.Button btn_Month;
		private System.Windows.Forms.ComboBox comboBoxMonth;
		private System.Windows.Forms.DateTimePicker dateTimePicker1;
		private System.Windows.Forms.ComboBox comboBoxYear;
		private System.Windows.Forms.Label lb_TemplatePath;
		private System.Windows.Forms.Label lb_OutputPath;
	}
}

