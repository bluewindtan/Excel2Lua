using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel2Lua
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Form1 formActivity = new Form1(LuaType.Activity);
			formActivity.ShowDialog();
		}

		private void button2_Click(object sender, EventArgs e)
		{
			Form1 formActivity = new Form1(LuaType.Packet);
			formActivity.ShowDialog();
		}

		private void button3_Click(object sender, EventArgs e)
		{
			Form1 formActivity = new Form1(LuaType.Box);
			formActivity.ShowDialog();
		}
	}
}
