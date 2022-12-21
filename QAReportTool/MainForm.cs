using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QAReportTool
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			this.Visible = false;
			FormCutFresh _frm = new FormCutFresh();
			_frm.ShowDialog();
			this.Visible = true;
		}

		private void button2_Click(object sender, EventArgs e)
		{
			this.Visible = false;
			FormFreshPartsHalal _frm = new FormFreshPartsHalal();
			_frm.ShowDialog();
			this.Visible = true;
		}

		private void button3_Click(object sender, EventArgs e)
		{
			this.Visible = false;
			FormFreshPartsNonHalal _frm = new FormFreshPartsNonHalal();
			_frm.ShowDialog();
			this.Visible = true;
		}

		private void button4_Click(object sender, EventArgs e)
		{
			this.Visible = false;
			FormFreshWholeChicken _frm = new FormFreshWholeChicken();
			_frm.ShowDialog();
			this.Visible = true;
		}

		private void button5_Click(object sender, EventArgs e)
		{
			this.Visible = false;
			FormFrozenThawed _frm = new FormFrozenThawed();
			_frm.ShowDialog();
			this.Visible = true;
		}
	}
}
