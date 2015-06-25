using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel2Lua;

namespace Excel2Lua
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			FolderBrowserDialog dlg = new FolderBrowserDialog();
			dlg.RootFolder = Environment.SpecialFolder.MyComputer;
			dlg.SelectedPath = Environment.CurrentDirectory;
			dlg.Description = "请选择目录";
			if (DialogResult.OK == dlg.ShowDialog())
			{
				string sDirectory = dlg.SelectedPath;
				ProcessDirectory(sDirectory);
				MessageBox.Show("Process finished！OK！");
			}
		}

		private void ProcessDirectory(string sDirectory)
		{
			// 判断是否目录 
			if (Directory.Exists(sDirectory))
			{
				DirectoryInfo dirInfo = new DirectoryInfo(sDirectory);
				foreach (FileSystemInfo fsInfo in dirInfo.GetFileSystemInfos())
				{
					if (fsInfo is FileInfo)
					{
						FileInfo fi = fsInfo as FileInfo;
						try
						{
							E2LDataMgr.ConvertE2L(fi.DirectoryName, fi.Name);
						}
						catch (System.Exception ex)
						{
							MessageBox.Show("Failed to convert excel '" + fi.Name + "'.\r\nError is " + ex.Message);
							continue;
						}
					}
					else if (fsInfo is DirectoryInfo)
					{
						ProcessDirectory(fsInfo.FullName);
					}
				}
			}
		}

	}
}
