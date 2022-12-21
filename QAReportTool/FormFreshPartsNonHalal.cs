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
	public partial class FormFreshPartsNonHalal : Form
	{
		static string DatabaseCompanyPrefix = ConfigurationManager.AppSettings["DatabaseCompanyPrefix"].ToString();
		string AppSettingDateFormat = ConfigurationManager.AppSettings["DateFormat"].ToString();
		string FreshCutReportTemplatePath = ConfigurationManager.AppSettings["FreshPartsNonHalalReportTemplatePath"].ToString();
		string OutputReportTemplatePath = ConfigurationManager.AppSettings["OutputPath"].ToString();
		public FormFreshPartsNonHalal()
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
							WHEN aa.No_ LIKE 'NP01' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP02' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP03' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP06' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP07' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP08' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP10' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP12' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP14' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP15' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'
							WHEN aa.No_ LIKE 'NP16' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal'

							End as 'Report Name',
 
							  case
							  when aa.[No_]  like 'NP%' THEN 'NH Fresh'
							END as 'Product Type'
							,aa.[No_] 'SKU',aa.[Description],
							case
							WHEN aa.No_ LIKE 'NP01' THEN 'a)Pasar Fresh Chicken Feet '
							WHEN aa.No_ LIKE 'NP02' THEN 'b)Pasar Fresh Chicken Wing '
							WHEN aa.No_ LIKE 'NP03' THEN 'c)Pasar Fresh Chicken Boneless Breast '
							WHEN aa.No_ LIKE 'NP06' THEN 'd)Pasar Fresh Chicken Leg'
							WHEN aa.No_ LIKE 'NP07' THEN 'e)Pasar Fresh Chicken Drumstick '
							WHEN aa.No_ LIKE 'NP08' THEN 'f)Pasar Fresh Chicken Thigh '
							WHEN aa.No_ LIKE 'NP10' THEN 'g)Pasar Fresh Chicken Fillet '
							WHEN aa.No_ LIKE 'NP12' THEN 'h)Pasar Fresh Chicken Minced '
							WHEN aa.No_ LIKE 'NP14' THEN 'i)Pasar Fresh Chicken Gizzard '
							WHEN aa.No_ LIKE 'NP15' THEN 'j)Pasar Fresh Chicken Bone '
							WHEN aa.No_ LIKE 'NP16' THEN 'k)Pasar Fresh Chicken Liver '
							END as 'QA Description',
							Case
							WHEN aa.No_ LIKE 'NP01' THEN 'a)221789'
							WHEN aa.No_ LIKE 'NP02' THEN 'b)234807'
							WHEN aa.No_ LIKE 'NP03' THEN 'c)234849'
							WHEN aa.No_ LIKE 'NP06' THEN 'd)234831'
							WHEN aa.No_ LIKE 'NP07' THEN 'e)234815'
							WHEN aa.No_ LIKE 'NP08' THEN 'f)234823'
							WHEN aa.No_ LIKE 'NP10' THEN 'g)234857'
							WHEN aa.No_ LIKE 'NP12' THEN 'h)234914'
							WHEN aa.No_ LIKE 'NP14' THEN 'i)234881'
							WHEN aa.No_ LIKE 'NP15' THEN 'j)212209'
							WHEN aa.No_ LIKE 'NP16' THEN 'k)234873'
							End as 'QA Code',
							Case
							WHEN aa.No_ LIKE 'NP01' THEN 'a)'
							WHEN aa.No_ LIKE 'NP02' THEN 'b)'
							WHEN aa.No_ LIKE 'NP03' THEN 'c)'
							WHEN aa.No_ LIKE 'NP06' THEN 'd)'
							WHEN aa.No_ LIKE 'NP07' THEN 'e)'
							WHEN aa.No_ LIKE 'NP08' THEN 'f)'
							WHEN aa.No_ LIKE 'NP10' THEN 'g)'
							WHEN aa.No_ LIKE 'NP12' THEN 'h)'
							WHEN aa.No_ LIKE 'NP14' THEN 'i)'
							WHEN aa.No_ LIKE 'NP15' THEN 'j)'
							WHEN aa.No_ LIKE 'NP16' THEN 'k)'
							End as 'OrderQtyAlphabet'
									,sum([Qty (Unit)]) 'Order Quantity',[Qty (Unit) UOM]
							  FROM [KEESONG].[dbo].[" + DatabaseCompanyPrefix + @"Sales Invoice Line] aa  with (NOLOCK)
							  left join [" + DatabaseCompanyPrefix + @"Item] cc with (NOLOCK) on aa.No_ = cc.No_
								left join [KEESONG].[dbo].[" + DatabaseCompanyPrefix + @"Item Unit of Measure] dd with (NOLOCK) on dd.[Item No_] = cc.No_ and dd.Code = [Qty (Unit) UOM]
							  where  [Shipment Date] = @ReportDate
							  and [Bill-to Customer No_] = 'M-N00050'
							  group by  aa.[No_],aa.[Description],[Qty (Unit) UOM]

							  ) aaa where aaa.[Product Type] <> 'NULL'
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
			if (!xxx.Contains("NP01"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP01", "鲜鸡脚    FRESH CHKN FEET", "a)Pasar Fresh Chicken Feet ", "a)221789", "a)", 0, "");

			if (!xxx.Contains("NP02"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP02", "鲜鸡翅膀 FRESH CHKN WING", "b)Pasar Fresh Chicken Wing ", "b)234807", "b)", 0, "");

			if (!xxx.Contains("NP03"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP03", "鲜无骨胸 FRESH B/L CHKN BREAST", "c)Pasar Fresh Chicken Boneless Breast ", "c)234849", "c)", 0, "");

			if (!xxx.Contains("NP06"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP06", "鲜有骨腿 FRESH B/IN CHKN LEG", "d)Pasar Fresh Chicken Leg", "d)234831", "d)", 0, "");

			if (!xxx.Contains("NP07"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP07", "鲜腿下部 FRESH CHKN DRUMSTICK", "e)Pasar Fresh Chicken Drumstick ", "e)234815", "e)", 0, "");

			if (!xxx.Contains("NP08"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP08", "鲜腿上部 FRESH CHKN THIGH", "f)Pasar Fresh Chicken Thigh ", "f)234823", "f)", 0, "");

			if (!xxx.Contains("NP10"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP10", "鲜鸡柳    FRESH CHKN FILLET", "g)Pasar Fresh Chicken Fillet ", "g)234857", "g)", 0, "");

			if (!xxx.Contains("NP12"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP12", "鲜搅碎肉 FRESH MINCED CHKN", "h)Pasar Fresh Chicken Minced ", "h)234914", "h)", 0, "");

			if (!xxx.Contains("NP14"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP14", "鸡腱 CHKN GIZZARD", "i)Pasar Fresh Chicken Gizzard ", "i)234881", "i)", 0, "");

			if (!xxx.Contains("NP15"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP15", "鸡骨 CHKN BONE", "j)Pasar Fresh Chicken Bone ", "j)212209", "j)", 0, "");

			if (!xxx.Contains("NP16"))
			rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Non-Halal", "NH Fresh", "NP16", "鸡肝 CHKN LIVER", "k)Pasar Fresh Chicken Liver ", "k)234873", "k)", 0, "");

			rInputDataTable.DefaultView.Sort = "OrderQtyAlphabet";
			rInputDataTable = rInputDataTable.DefaultView.ToTable();

			return rInputDataTable;


		}

	}
}
