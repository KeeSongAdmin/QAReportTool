using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QAReportTool
{
	public partial class FormCutFresh : Form
	{
		static string DatabaseCompanyPrefix = ConfigurationManager.AppSettings["DatabaseCompanyPrefix"].ToString();
		string AppSettingDateFormat = ConfigurationManager.AppSettings["DateFormat"].ToString();
		string FreshCutReportTemplatePath = ConfigurationManager.AppSettings["FreshCutReportTemplatePath"].ToString(); 
		string OutputReportTemplatePath = ConfigurationManager.AppSettings["OutputPath"].ToString(); 
		public FormCutFresh()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			dateTimePicker1.Format = DateTimePickerFormat.Custom;
			dateTimePicker1.CustomFormat = "yyyy-MM-dd";
			#region combobox
			Dictionary<int, string> comboSource = new Dictionary<int, string>();
			comboSource.Add(0, "Please Select");
			comboSource.Add(1, "January");
			comboSource.Add(2, "February");
			comboSource.Add(3, "March");
			comboSource.Add(4, "April");
			comboSource.Add(5, "May");
			comboSource.Add(6, "June");
			comboSource.Add(7, "July");
			comboSource.Add(8, "August");
			comboSource.Add(9, "September");
			comboSource.Add(10, "October");
			comboSource.Add(11, "November");
			comboSource.Add(12, "December");
			comboBoxMonth.DataSource = new BindingSource(comboSource, null);
			comboBoxMonth.DisplayMember = "Value";
			comboBoxMonth.ValueMember = "Key";
			
			Dictionary<int, string> comboSource2 = new Dictionary<int, string>();
			comboSource2.Add(0, "Please Select");
			comboSource2.Add(DateTime.Now.Year - 1, (DateTime.Now.Year - 1).ToString());
			comboSource2.Add(DateTime.Now.Year, (DateTime.Now.Year).ToString());
			comboSource2.Add(DateTime.Now.Year + 1, (DateTime.Now.Year + 1).ToString());
			comboSource2.Add(DateTime.Now.Year + 2, (DateTime.Now.Year + 2).ToString());
			comboBoxYear.DataSource = new BindingSource(comboSource2, null);
			comboBoxYear.DisplayMember = "Value";
			comboBoxYear.ValueMember = "Key";
			#endregion

			lb_TemplatePath.Text = "Template Path : "+FreshCutReportTemplatePath;
			lb_OutputPath.Text = "Output Path : "+FreshCutReportTemplatePath;

		}

		private string GenerateSQLQuery(String rUserSelectedDate)
		{
			string query = @"DECLARE @ReportDate date = '" + rUserSelectedDate + @"'

								select * from
								(
								SELECT CASE 
								WHEN aa.No_ LIKE 'ZHC12' THEN N'STOCK QC FORM 货品品质检查表_Fresh Cut'
								WHEN aa.No_ LIKE 'ZCC12' THEN N'STOCK QC FORM 货品品质检查表_Fresh Cut'
								End as 'Report Name',
								CASE
								WHEN aa.[No_]  like 'ZHC12' THEN 'Halal Cut 12'
								WHEN aa.[No_]  like 'ZCC12' THEN 'NH Cut 12'
								END as 'Product Type'
								,aa.[No_] 'SKU',aa.[Description],
								CASE
								WHEN aa.No_ LIKE 'ZHC12' THEN 'a)Pasar Fresh Cut Chicken (12 Portions) - Halal'
								WHEN aa.No_ LIKE 'ZCC12' THEN 'b)Pasar Fresh Cut Chicken (12 Portions)'
								END as 'QA Description',
								CASE
								WHEN aa.No_ LIKE 'ZHC12' THEN 'a)10787302'
								WHEN aa.No_ LIKE 'ZCC12' THEN 'b)13102624'
								End as 'QA Code',
								Case
								WHEN aa.No_ LIKE 'ZHC12' THEN 'a)'
								WHEN aa.No_ LIKE 'ZCC12' THEN 'b)'
								End as 'OrderQtyAlphabet'
										,SUM([Qty (Unit)]) 'Order Quantity',[Qty (Unit) UOM]
								FROM [KEESONG].[dbo].[" + DatabaseCompanyPrefix + @"Sales Invoice Line] aa  with (NOLOCK)
								left join [" + DatabaseCompanyPrefix + @"Item] cc with (NOLOCK) on aa.No_ = cc.No_
								left join [KEESONG].[dbo].[" + DatabaseCompanyPrefix + @"Item Unit of Measure] dd with (NOLOCK) on dd.[Item No_] = cc.No_ and dd.Code = [Qty (Unit) UOM]
								WHERE  [Shipment Date] = @ReportDate AND [Bill-to Customer No_] = 'M-N00050'
								GROUP BY aa.[No_],aa.[Description],[Qty (Unit) UOM]
								)aaa where aaa.[Product Type] <> 'NULL'
								order by [Report Name],[Product Type],SKU	";

			return query;

		}
		public void btn_Day_Click(object sender, EventArgs e)
		{
			btn_Day.Enabled = false;
			try
			{
				String userSelectedDateString = dateTimePicker1.Value.ToString(AppSettingDateFormat);
				DateTime userSelectedDate = dateTimePicker1.Value;
				DataTable dtTemp = new DataTable();
				ClassDB cls_db = new ClassDB();
				int userSelectedYear = userSelectedDate.Year;
				int userSelectedMonth = userSelectedDate.Month;
				String userSelectedMMonth = userSelectedDate.ToString("MMMM");

				if (userSelectedDateString == "")
					MessageBox.Show("Please Select Date");
				else
				{
					dtTemp = new DataTable();
					dtTemp = cls_db.SelectQueryNoLock(GenerateSQLQuery(userSelectedDateString), ConfigurationManager.ConnectionStrings["DefaultDBconn"].ConnectionString);
					dtTemp = FillEmptySKU(dtTemp);
					String fileName = Path.GetFileNameWithoutExtension(FreshCutReportTemplatePath).Replace("Template-", "") + "(" + userSelectedMMonth + " " + userSelectedYear + ")";
					String fileExt = Path.GetExtension(FreshCutReportTemplatePath);
					String filePath = Path.GetDirectoryName(FreshCutReportTemplatePath);
					String savepath = OutputReportTemplatePath + @"\" + fileName + fileExt;
					String appendFilePath = File.Exists(savepath) ? savepath : FreshCutReportTemplatePath;
					lb_OutputPath.Text = appendFilePath;
					new WriteExcel().AppendExcelMultipleSheetBySheetIndex(appendFilePath, savepath, dtTemp, userSelectedDate.Day);

				}
				btn_Day.Enabled = true;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Please look for IT . program encounter error \r\n" + ex.Message.ToString() + "\r\n" + ex.InnerException.ToString());
			}
			finally
			{

			}
		}

		public void btn_Month_Click(object sender, EventArgs e)
		{
			btn_Month.Enabled = false;
			try
			{
				String userSelectedDate = dateTimePicker1.Value.ToString(AppSettingDateFormat);
				List<DataTable> resultTableList = new List<DataTable>();
				DataTable dtTemp = new DataTable();
				ClassDB cls_db = new ClassDB();
				int userSelectedYear = int.Parse(comboBoxYear.SelectedValue.ToString());
				int userSelectedMonth = int.Parse(comboBoxMonth.SelectedValue.ToString());
				String userSelectedMMonth = comboBoxMonth.Text.ToString();

				if (true)//month
				{
					if (userSelectedMonth == 0 || userSelectedYear == 0)
						MessageBox.Show("Please Select Valid Month & Year");
					else
					{
						int days = DateTime.DaysInMonth(userSelectedYear, userSelectedMonth);
						for (int i = 0; i < days; i++)
						{
							DateTime dateTemp = new DateTime(userSelectedYear, userSelectedMonth, 1);
							String forloopDateString = dateTemp.AddDays(i).ToString(AppSettingDateFormat);
							dtTemp = new DataTable();
							dtTemp = cls_db.SelectQueryNoLock(GenerateSQLQuery(forloopDateString), ConfigurationManager.ConnectionStrings["DefaultDBconn"].ConnectionString);
							dtTemp = FillEmptySKU(dtTemp);
							resultTableList.Add(dtTemp);
						}
						String fileName = Path.GetFileNameWithoutExtension(FreshCutReportTemplatePath).Replace("Template-", "")+"("+ userSelectedMMonth +" "+ userSelectedYear + ")";
						String fileExt = Path.GetExtension(FreshCutReportTemplatePath);
						String filePath = Path.GetDirectoryName(FreshCutReportTemplatePath);
						String savepath = OutputReportTemplatePath + @"\" + fileName + fileExt;

						lb_OutputPath.Text = savepath;
						new WriteExcel().AppendExcelMultipleSheet(FreshCutReportTemplatePath, savepath, resultTableList);
					}
				}
				
				btn_Month.Enabled = true;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Please look for IT . program encounter error \r\n" + ex.Message.ToString() + "\r\n" + ex.InnerException.ToString());
			}
			finally
			{

			}


		}

		private void comboBoxYear_SelectedIndexChanged(object sender, EventArgs e)
		{
			GetOutputPath();
		}

		private void comboBoxMonth_SelectedIndexChanged(object sender, EventArgs e)
		{
			GetOutputPath();
		}

		private void GetOutputPath()
		{
			int userSelectedYear = comboBoxYear.SelectedIndex > 0 ? int.Parse(comboBoxYear.SelectedValue.ToString()) : 0;
			int userSelectedMonth = comboBoxMonth.SelectedIndex > 0 ? int.Parse(comboBoxMonth.SelectedValue.ToString()) : 0;
			String userSelectedMMonth = comboBoxMonth.Text.ToString();

			String fileName = Path.GetFileNameWithoutExtension(FreshCutReportTemplatePath).Replace("Template-", "") + "(" + userSelectedMMonth + " " + userSelectedYear + ")";
			String fileExt = Path.GetExtension(FreshCutReportTemplatePath);
			String filePath = Path.GetDirectoryName(FreshCutReportTemplatePath);
			String savepath = OutputReportTemplatePath + @"\" + fileName + fileExt;
			lb_OutputPath.Text = savepath;
		}

		private void btn_Back_Click(object sender, EventArgs e)
		{
			this.Close();
		}


		private DataTable FillEmptySKU(DataTable rInputDataTable)
		{
			List<String> xxx = rInputDataTable.AsEnumerable().Select(x => x["SKU"].ToString()).ToList();
			if (!xxx.Contains("ZHC12"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Cut", "Halal Cut 12", "ZHC12", "PASAR FRESH CUT CHKN (H)-12 portions", "a)Pasar Fresh Cut Chicken (12 Portions) - Halal", "a)10787302", "a)", 0, "");

			if (!xxx.Contains("ZCC12"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Cut", "NH Cut 12", "ZCC12", "PASAR FRESH CUT CHKN (NH)-12PORTION", "b)Pasar Fresh Cut Chicken (12 Portions)", "b)13102624", "b)", 0, "");
			 

			rInputDataTable.DefaultView.Sort = "OrderQtyAlphabet";
			rInputDataTable = rInputDataTable.DefaultView.ToTable();

			return rInputDataTable;


		}

	}
}
