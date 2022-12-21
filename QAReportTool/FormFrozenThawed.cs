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
	public partial class FormFrozenThawed : Form
	{
		static string DatabaseCompanyPrefix = ConfigurationManager.AppSettings["DatabaseCompanyPrefix"].ToString();
		string AppSettingDateFormat = ConfigurationManager.AppSettings["DateFormat"].ToString();
		string FreshCutReportTemplatePath = ConfigurationManager.AppSettings["FrozenThawedReportTemplatePath"].ToString();
		string OutputReportTemplatePath = ConfigurationManager.AppSettings["OutputPath"].ToString();
		public FormFrozenThawed()
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
								SELECT  
								case
							WHEN aa.No_ LIKE 'VH01' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH02' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH03' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH04' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH05' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH06' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH07' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH08' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'
							WHEN aa.No_ LIKE 'VH09' THEN N'STOCK QC FORM 货品品质检查表_Frozen Thawed'

							End as 'Report Name',
 
								case
								when aa.[No_]  like 'VH0%' THEN 'FT 300G'
							END as 'Product Type'
							,aa.[No_] 'SKU',aa.[Description],
							case
							WHEN aa.No_ LIKE 'VH01' THEN 'a)Pasar Frozen Thawed Chicken Fillet (300g) - Halal'
							WHEN aa.No_ LIKE 'VH02' THEN 'b)Pasar Frozen Thawed Chicken Boneless Breast (300g) - Halal'
							WHEN aa.No_ LIKE 'VH03' THEN 'c)Pasar Frozen Thawed Chicken Boneless Leg (300g) - Halal'
							WHEN aa.No_ LIKE 'VH04' THEN 'd)Pasar Frozen Thawed Chicken Wing (300g) - Halal'
							WHEN aa.No_ LIKE 'VH05' THEN 'e)Pasar Frozen Thawed Chicken Drumstick (300g) - Halal'
							WHEN aa.No_ LIKE 'VH06' THEN 'f)Pasar Frozen Thawed Chicken Thigh (300g) - Halal'
							WHEN aa.No_ LIKE 'VH07' THEN 'g)Pasar Frozen Thawed Chicken Drumette (300g) - Halal'
							WHEN aa.No_ LIKE 'VH08' THEN 'h)Pasar Frozen Thawed Chicken Mid Joint Wing (300g) - Halal'
							WHEN aa.No_ LIKE 'VH09' THEN 'i)Pasar Frozen Thawed Chicken Pieces (500g) - Halal '
							END as 'QA Description',
							Case
							WHEN aa.No_ LIKE 'VH01' THEN 'a)10771829'
							WHEN aa.No_ LIKE 'VH02' THEN 'b)10771861'
							WHEN aa.No_ LIKE 'VH03' THEN 'c)10771896'
							WHEN aa.No_ LIKE 'VH04' THEN 'd)10771757'
							WHEN aa.No_ LIKE 'VH05' THEN 'e)10771845'
							WHEN aa.No_ LIKE 'VH06' THEN 'f)10771773'
							WHEN aa.No_ LIKE 'VH07' THEN 'g)10771730'
							WHEN aa.No_ LIKE 'VH08' THEN 'h)10771802'
							WHEN aa.No_ LIKE 'VH09' THEN 'i)13218511'
							End as 'QA Code',
									Case
							WHEN aa.No_ LIKE 'VH01' THEN 'a)'
							WHEN aa.No_ LIKE 'VH02' THEN 'b)'
							WHEN aa.No_ LIKE 'VH03' THEN 'c)'
							WHEN aa.No_ LIKE 'VH04' THEN 'd)'
							WHEN aa.No_ LIKE 'VH05' THEN 'e)'
							WHEN aa.No_ LIKE 'VH06' THEN 'f)'
							WHEN aa.No_ LIKE 'VH07' THEN 'g)'
							WHEN aa.No_ LIKE 'VH08' THEN 'h)'
							WHEN aa.No_ LIKE 'VH09' THEN 'i)'
							End as 'OrderQtyAlphabet'
									,sum([Qty (Unit)]) 'Order Quantity',[Qty (Unit) UOM]
							FROM [KEESONG].[dbo].[" + DatabaseCompanyPrefix + @"Sales Invoice Line] aa  with (NOLOCK)
							left join [" + DatabaseCompanyPrefix + @"Item] cc with (NOLOCK) on aa.No_ = cc.No_
							left join [KEESONG].[dbo].[" + DatabaseCompanyPrefix + @"Item Unit of Measure] dd with (NOLOCK) on dd.[Item No_] = cc.No_ and dd.Code = [Qty (Unit) UOM]
								where  [Shipment Date] = @ReportDate
								and [Bill-to Customer No_] = 'M-N00050'
								group by  aa.[No_],aa.[Description],[Qty (Unit) UOM]

								) aaa where aaa.[Product Type] <> 'NULL'
								order by OrderQtyAlphabet	";

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
			if (!xxx.Contains("VH01"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH01", "鸡柳 (300g) (H) FT CHKN FILLET 300G", "a)Pasar Frozen Thawed Chicken Fillet (300g) - Halal", "a)10771829", "a)", 0, "");

			if (!xxx.Contains("VH02"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH02", "无骨胸(300g) (H) FT CHKN B/L BREAST 300G", "b)Pasar Frozen Thawed Chicken Boneless Breast (300g) - Halal", "b)10771861", "b)", 0, "");

			if (!xxx.Contains("VH03"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH03", "无骨腿(300g) (H) FT CHKN B/L LEG 300G", "c)Pasar Frozen Thawed Chicken Boneless Leg (300g) - Halal", "c)10771896", "c)", 0, "");

			if (!xxx.Contains("VH04"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH04", "翅膀(300g) (H) FT CHKN WING 300G", "d)Pasar Frozen Thawed Chicken Wing (300g) - Halal", "d)10771757", "d)", 0, "");
			if (!xxx.Contains("VH05"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH05", "腿下部(300g) (H) FT CHKN DRUMSTICK 300G", "e)Pasar Frozen Thawed Chicken Drumstick (300g) - Halal", "e)10771845", "e)", 0, "");
			if (!xxx.Contains("VH06"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH06", "腿上部(300g) (H) FT CHKN THIGH 300G", "f)Pasar Frozen Thawed Chicken Thigh (300g) - Halal", "f)10771773", "f)", 0, "");
			if (!xxx.Contains("VH07"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH07", "翅上部(300g) (H) FT CHKN DRUMMETTE 300G", "g)Pasar Frozen Thawed Chicken Drumette (300g) - Halal", "g)10771730", "g)", 0, "");
			if (!xxx.Contains("VH08"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH08", "翅中部(300g) (H) FT CHKN MID JOINT WING 300G", "h)Pasar Frozen Thawed Chicken Mid Joint Wing (300g) - Halal", "h)10771802", "h)", 0, "");
			if (!xxx.Contains("VH09"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Frozen Thawed", "FT 300G", "VH09", "鸡块 FROZEN THAWED CHICKEN PIECES", "i)Pasar Frozen Thawed Chicken Pieces (500g) - Halal ", "i)13218511", "i)", 0, "");



			rInputDataTable.DefaultView.Sort = "OrderQtyAlphabet";
			rInputDataTable = rInputDataTable.DefaultView.ToTable();

			return rInputDataTable;


		}

	}
}
