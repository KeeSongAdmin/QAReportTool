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
	public partial class FormFreshWholeChicken : Form
	{
		static string DatabaseCompanyPrefix = ConfigurationManager.AppSettings["DatabaseCompanyPrefix"].ToString();
		string AppSettingDateFormat = ConfigurationManager.AppSettings["DateFormat"].ToString();
		string FreshCutReportTemplatePath = ConfigurationManager.AppSettings["FreshWholeChickenReportTemplatePath"].ToString();
		string OutputReportTemplatePath = ConfigurationManager.AppSettings["OutputPath"].ToString();
		public FormFreshWholeChicken()
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

							WHEN aa.No_ LIKE 'ZC10' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'ZH10' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'ZC14' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'ZH14' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'ZC71' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'ZH71' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'FL20' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'HFL21' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'LB01' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'ZC89' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'
							WHEN aa.No_ LIKE 'ZH89' THEN N'STOCK QC FORM 货品品质检查表_Fresh Whole Chicken'

							End as 'Report Name',
 
								case
								when aa.[No_]  like 'FL20' THEN 'Kampong Chicken (NH)'
								when aa.[No_]  like 'HFL21' THEN 'Kampong Chicken (H)'
								when aa.[No_]  like 'ZC10' THEN N'八宝-NH (1kg)'
								when aa.[No_]  like 'ZH10' THEN N'八宝-H (1kg)'
								when aa.[No_]  like 'ZC14' THEN N'农宝-NH (1.4kg)'
								when aa.[No_]  like 'ZH14' THEN N'农宝-H (1.4kg)'
								when aa.[No_]  like 'ZC71' THEN N'珍宝-NH (1.7kg)'
								when aa.[No_]  like 'ZH71' THEN N'珍宝-H (1.7kg)'
								when aa.[No_]  like 'ZC89' THEN N'鲜鸡-NH (800g)'
								when aa.[No_]  like 'ZH89' THEN N'鲜鸡-H (800g)'
								when aa.[No_]  like 'LB01' THEN 'Black Chicken'
							END as 'Product Type'
							,aa.[No_] 'SKU',aa.[Description],
							case
							WHEN aa.No_ LIKE 'ZC10' THEN 'a)Pasar Fresh Chicken 1kg'
							WHEN aa.No_ LIKE 'ZH10' THEN 'b)Pasar Fresh Chicken 1kg - Halal'
							WHEN aa.No_ LIKE 'ZC14' THEN 'c)Pasar Farm Fresh Chicken 1.4kg'
							WHEN aa.No_ LIKE 'ZH14' THEN 'd)Pasar Farm Fresh Chicken 1.4kg - Halal'
							WHEN aa.No_ LIKE 'ZC71' THEN 'e)Fresh Jumbo Fresh Chicken 1.7kg'
							WHEN aa.No_ LIKE 'ZH71' THEN 'f)Fresh Jumbo Fresh Chicken 1.7kg - Halal'
							WHEN aa.No_ LIKE 'FL20' THEN 'g)Pasar Fresh Kampong Chicken'
							WHEN aa.No_ LIKE 'HFL21' THEN 'h)Pasar Fresh Kampong Chicken - Halal'
							WHEN aa.No_ LIKE 'LB01' THEN 'i)Pasar Fresh Black Chicken (Note: In Tray Pack Only)'
							WHEN aa.No_ LIKE 'ZC89' THEN 'j)Pasar Fresh Chicken 800g (Note: In Tray Pack Only)'
							WHEN aa.No_ LIKE 'ZH89' THEN 'k)Pasar Fresh Chicken 800g - Halal (Note: In Tray Pack Only)'
							END as 'QA Description',
							Case
							WHEN aa.No_ LIKE 'ZC10' THEN 'a)481408'
							WHEN aa.No_ LIKE 'ZH10' THEN 'b)10001736'
							WHEN aa.No_ LIKE 'ZC14' THEN 'c)410118'
							WHEN aa.No_ LIKE 'ZH14' THEN 'd)212051'
							WHEN aa.No_ LIKE 'ZC71' THEN 'e)10901494'
							WHEN aa.No_ LIKE 'ZH71' THEN 'f)10901398'
							WHEN aa.No_ LIKE 'FL20' THEN 'g)10564811'
							WHEN aa.No_ LIKE 'HFL21' THEN 'h)10564774'
							WHEN aa.No_ LIKE 'LB01' THEN 'i)234980'
							WHEN aa.No_ LIKE 'ZC89' THEN 'j)10992920'
							WHEN aa.No_ LIKE 'ZH89' THEN 'k)10992920'
							End as 'QA Code',
							Case
							WHEN aa.No_ LIKE 'ZC10' THEN 'a)'
							WHEN aa.No_ LIKE 'ZH10' THEN 'b)'
							WHEN aa.No_ LIKE 'ZC14' THEN 'c)'
							WHEN aa.No_ LIKE 'ZH14' THEN 'd)'
							WHEN aa.No_ LIKE 'ZC71' THEN 'e)'
							WHEN aa.No_ LIKE 'ZH71' THEN 'f)'
							WHEN aa.No_ LIKE 'FL20' THEN 'g)'
							WHEN aa.No_ LIKE 'HFL21' THEN 'h)'
							WHEN aa.No_ LIKE 'LB01' THEN 'i)'
							WHEN aa.No_ LIKE 'ZC89' THEN 'j)'
							WHEN aa.No_ LIKE 'ZH89' THEN 'k)'
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

		private DataTable FillEmptySKU(DataTable rInputDataTable)
		{
			List<String> xxx = rInputDataTable.AsEnumerable().Select(x => x["SKU"].ToString()).ToList();
			if(!xxx.Contains("ZC10"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"八宝 - NH(1kg)",		"ZC10",		"鲜鸡(巴) 1.0Kg - 1.3Kg FRESH CHKN(S)",			"a)Pasar Fresh Chicken 1kg",									"a)481408",     "a)", 0, "");

			if (!xxx.Contains("ZH10"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"八宝 - H(1kg)",			"ZH10",		"马来鸡(巴) 1.0Kg - 1.3Kg HALAL CHKN",			"b)Pasar Fresh Chicken 1kg - Halal",							"b)10001736",   "b)", 0, "");

			if (!xxx.Contains("ZC14"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"农宝 - NH(1.4kg)",		"ZC14",		"农宝鲜鸡1.4Kg - 1.6Kg FARM FRESH CHKN",			"c)Pasar Farm Fresh Chicken 1.4kg",								"c)410118",     "c)", 0, "");

			if (!xxx.Contains("ZH14"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"农宝 - H(1.4kg)",		"ZH14",		"马来农宝鲜鸡1.4Kg - 1.6Kg H FARM FRESH CHKN",	"d)Pasar Farm Fresh Chicken 1.4kg - Halal",						"d)212051",     "d)", 0, "");

			if (!xxx.Contains("ZC71"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"珍宝 - NH(1.7kg)",		"ZC71",		"珍宝鲜鸡 1.7Kg - 2.0Kg JUMBO FRESH CHKN",		"e)Fresh Jumbo Fresh Chicken 1.7kg",							"e)10901494",   "e)", 0, "");

			if (!xxx.Contains("ZH71"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"珍宝 - H(1.7kg)",		"ZH71",		"马来珍宝1.7Kg - 2.0Kg HALAL JUMBO FRESH CHKN",	"f)Fresh Jumbo Fresh Chicken 1.7kg - Halal",					"f)10901398",	"f)", 0, "");

			if (!xxx.Contains("FL20"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"Kampong Chicken(NH)",	"FL20",		"包装甘榜鸡900g-1.2Kg PK KAMPONG CHKN",			"g)Pasar Fresh Kampong Chicken",								"g)10564811",	"g)", 0, "");

			if (!xxx.Contains("HFL21"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"Kampong Chicken(H)",	"HFL21",	"HALAL KAMPONG FRESH CHKN 900GM - 1.2KG",		"h)Pasar Fresh Kampong Chicken -Halal",							"h)10564774",	"h)", 0, "");

			if (!xxx.Contains("LB01"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"Black Chicken",		"LB01",		"鲜黑鸡 FRESH BLACK CHKN",						"i)Pasar Fresh Black Chicken(Note: In Tray Pack Only)",			"i)234980",		"i)", 0, "");

			if (!xxx.Contains("ZC89"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"鲜鸡 - NH(800g)",		"ZC89",		"鲜鸡 800g - 900g SP PASAR FRESH CHKN",			"j)Pasar Fresh Chicken 800g(Note: In Tray Pack Only)",			"j)10992920",	"j)", 0, "");

			if (!xxx.Contains("ZH89"))
				rInputDataTable.Rows.Add("STOCK QC FORM 货品品质检查表_Fresh Whole Chicken",	"鲜鸡 - H(800g)",		"ZH89",		"马来鸡 800g - 900g SP PASAR HALAL CHKN",		"k)Pasar Fresh Chicken 800g - Halal(Note: In Tray Pack Only)",	"k)10992920",	"k)", 0, "");


			rInputDataTable.DefaultView.Sort = "OrderQtyAlphabet";
			rInputDataTable = rInputDataTable.DefaultView.ToTable();

			return rInputDataTable;
		
		
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
		 
	}
}
