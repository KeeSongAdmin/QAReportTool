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
	public partial class FormFreshPartsHalal : Form
	{
		static string DatabaseCompanyPrefix = ConfigurationManager.AppSettings["DatabaseCompanyPrefix"].ToString();
		string AppSettingDateFormat = ConfigurationManager.AppSettings["DateFormat"].ToString();
		string FreshCutReportTemplatePath = ConfigurationManager.AppSettings["FreshPartsHalalReportTemplatePath"].ToString();
		string OutputReportTemplatePath = ConfigurationManager.AppSettings["OutputPath"].ToString();
		public FormFreshPartsHalal()
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
							WHEN aa.No_ LIKE 'NH01' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH02' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH03' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH06' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH07' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH08' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH10' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH12' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH14' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH15' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							WHEN aa.No_ LIKE 'NH16' THEN N'STOCK QC FORM 货品品质检查表_Fresh Parts Halal'
							End as 'Report Name',
 
							  case 
							  when aa.[No_]  like 'NH%' THEN 'Halal Fresh'
							END as 'Product Type'
							,aa.[No_] 'SKU',aa.[Description],
							case
							WHEN aa.No_ LIKE 'NH01' THEN 'a)Pasar Fresh Chicken Feet - Halal'
							WHEN aa.No_ LIKE 'NH02' THEN 'b)Pasar Fresh Chicken Wing - Halal'
							WHEN aa.No_ LIKE 'NH03' THEN 'c)Pasar Fresh Chicken Boneless Breast - Halal'
							WHEN aa.No_ LIKE 'NH06' THEN 'd)Pasar Fresh Chicken Whole Leg - Halal'
							WHEN aa.No_ LIKE 'NH07' THEN 'e)Pasar Fresh Chicken Drumstick - Halal'
							WHEN aa.No_ LIKE 'NH08' THEN 'f)Pasar Fresh Chicken Thigh - Halal'
							WHEN aa.No_ LIKE 'NH10' THEN 'g)Pasar Fresh Chicken Fillet - Halal'
							WHEN aa.No_ LIKE 'NH12' THEN 'h)Pasar Fresh Chicken Minced - Halal'
							WHEN aa.No_ LIKE 'NH14' THEN 'i)Pasar Fresh Chicken Gizzard - Halal'
							WHEN aa.No_ LIKE 'NH15' THEN 'j)Pasar Fresh Chicken Bone - Halal'
							WHEN aa.No_ LIKE 'NH16' THEN 'k)Pasar Fresh Chicken Liver - Halal'
							END as 'QA Description',
							Case
							WHEN aa.No_ LIKE 'NH01' THEN 'a)10001867'
							WHEN aa.No_ LIKE 'NH02' THEN 'b)10006270'
							WHEN aa.No_ LIKE 'NH03' THEN 'c)10001840'
							WHEN aa.No_ LIKE 'NH06' THEN 'd)10001955'
							WHEN aa.No_ LIKE 'NH07' THEN 'e)10001859'
							WHEN aa.No_ LIKE 'NH08' THEN 'f)10005876'
							WHEN aa.No_ LIKE 'NH10' THEN 'g)10001891'
							WHEN aa.No_ LIKE 'NH12' THEN 'h)10002253'
							WHEN aa.No_ LIKE 'NH14' THEN 'i)10001904'
							WHEN aa.No_ LIKE 'NH15' THEN 'j)10064470'
							WHEN aa.No_ LIKE 'NH16' THEN 'k)10001971'
							End as 'QA Code',
							Case
							WHEN aa.No_ LIKE 'NH01' THEN 'a)'
							WHEN aa.No_ LIKE 'NH02' THEN 'b)'
							WHEN aa.No_ LIKE 'NH03' THEN 'c)'
							WHEN aa.No_ LIKE 'NH06' THEN 'd)'
							WHEN aa.No_ LIKE 'NH07' THEN 'e)'
							WHEN aa.No_ LIKE 'NH08' THEN 'f)'
							WHEN aa.No_ LIKE 'NH10' THEN 'g)'
							WHEN aa.No_ LIKE 'NH12' THEN 'h)'
							WHEN aa.No_ LIKE 'NH14' THEN 'i)'
							WHEN aa.No_ LIKE 'NH15' THEN 'j)'
							WHEN aa.No_ LIKE 'NH16' THEN 'k)'
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
			if (!xxx.Contains("NH01"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH01", "鲜鸡脚    FRESH CHKN FEET", "a)Pasar Fresh Chicken Feet - Halal", "a)10001867", "a)", 0, "");

			if (!xxx.Contains("NH02"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH02", "鲜鸡翅膀 FRESH CHKN WING", "b)Pasar Fresh Chicken Wing - Halal", "b)10006270", "b)", 0, "");

			if (!xxx.Contains("NH03"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH03", "鲜无骨胸 FRESH B/L CHKN BREAST", "c)Pasar Fresh Chicken Boneless Breast - Halal", "c)10001840", "c)", 0, "");

			if (!xxx.Contains("NH06"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH06", "鲜有骨腿 FRESH B/IN CHKN LEG", "d)Pasar Fresh Chicken Whole Leg - Halal", "d)10001955", "d)", 0, "");

			if (!xxx.Contains("NH07"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH07", "鲜腿下部 FRESH CHKN DRUMSTICK", "e)Pasar Fresh Chicken Drumstick - Halal", "e)10001859", "e)", 0, "");

			if (!xxx.Contains("NH08"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH08", "鲜腿上部 FRESH CHKN THIGH", "f)Pasar Fresh Chicken Thigh - Halal", "f)10005876", "f)", 0, "");

			if (!xxx.Contains("NH10"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH10", "鲜鸡柳    FRESH CHKN FILLET", "g)Pasar Fresh Chicken Fillet - Halal", "g)10001891", "g)", 0, "");

			if (!xxx.Contains("NH12"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH12", "鲜搅碎肉 FRESH MINCED CHKN", "h)Pasar Fresh Chicken Minced - Halal", "h)10002253", "h)", 0, "");

			if (!xxx.Contains("NH14"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH14", "鸡腱 CHKN GIZZARD", "i)Pasar Fresh Chicken Gizzard - Halal", "i)10001904", "i)", 0, "");

			if (!xxx.Contains("NH15"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH15", "鸡骨 CHKN BONE", "j)Pasar Fresh Chicken Bone - Halal", "j)10064470", "j)", 0, "");

			if (!xxx.Contains("NH16"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Parts Halal", "Halal Fresh", "NH16", "鸡肝 CHKN LIVER", "k)Pasar Fresh Chicken Liver - Halal", "k)10001971", "k)", 0, "");




			rInputDataTable.DefaultView.Sort = "OrderQtyAlphabet";
			rInputDataTable = rInputDataTable.DefaultView.ToTable();

			return rInputDataTable;


		}


	}
}
