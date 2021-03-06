﻿using System;
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
		private LuaType m_luaType = LuaType.Activity;
		public Form1(LuaType nType)
		{
			m_luaType = nType;
			InitializeComponent();
			this.Text = CustomDefine.TOOL_TEXT[(int)m_luaType];
		}

		private void button1_Click(object sender, EventArgs e)
		{
			bool bExcelSuffix = checkBox1.Checked;
			FolderBrowserDialog dlg = new FolderBrowserDialog();
			dlg.RootFolder = Environment.SpecialFolder.MyComputer;
			dlg.SelectedPath = Environment.CurrentDirectory;
			dlg.Description = "Please select a folder:";
			if (DialogResult.OK == dlg.ShowDialog())
			{
				LogMgr.Instance.OpenLog();
				string sDirectory = dlg.SelectedPath;
				if (ProcessDirectory(sDirectory, bExcelSuffix))
				{
					MessageBox.Show("Process finished！OK！");
				}
				LogMgr.Instance.CloseLog();
			}
		}

		private bool ProcessDirectory(string sDirectory, bool bExcelSuffix)
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
							if (bExcelSuffix && !fi.Name.EndsWith(".xls") && !fi.Name.EndsWith(".xlsx"))
							{
								continue;
							}
							else 
							{
								E2LDataMgr.Instance.ConvertE2L(fi.DirectoryName, fi.Name);
							}
						}
						catch (System.Exception ex)
						{
							string strMessage = "Failed to convert excel '" + fi.Name + "'.\r\nError is " + ex.Message;
							LogMgr.Instance.WriteLog(strMessage);
							strMessage += "\r\nDo you want to continue???";
							DialogResult dlgResult = MessageBox.Show(strMessage, "error", MessageBoxButtons.OKCancel);
							if (DialogResult.OK == dlgResult)
							{
								continue;
							}
							else
							{
								return false;
							}
						}
					}
					else if (fsInfo is DirectoryInfo)
					{
						ProcessDirectory(fsInfo.FullName, bExcelSuffix);
					}
				}
			}
			else
			{
				MessageBox.Show("Not exists " + sDirectory);
				return false;
			}

			return true;
		}

	}
}
