using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QAReportTool
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			string Mode = ConfigurationManager.AppSettings["Mode"].ToString();
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			if(Mode.ToUpper()=="SCHEDULER")
			{
				new FormCutFresh().btn_Day_Click(null, null);
				new FormFreshPartsHalal().btn_Day_Click(null, null);
				new FormFreshPartsNonHalal().btn_Day_Click(null, null);
				new FormFreshWholeChicken().btn_Day_Click(null, null);
				new FormFrozenThawed().btn_Day_Click(null, null);
			}
			else
				Application.Run(new MainForm());






		}
	}
}
