using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using MSExcelReader;

namespace Excel2Lua
{
	public class E2L_Box
	{
		string m_strFullExcelName = "";
		string m_strDirector = "";
		string m_strExcel = "";		

		ExcelReader m_excelReader = null;
		StreamWriter m_sWriter = null;

		List<cBoxInfo> m_listBoxInfo = new List<cBoxInfo>();

		public E2L_Box(string strDirector, string strExcel)
		{
			m_strFullExcelName = strDirector + "\\" + strExcel;
			m_strDirector = strDirector;		
			m_strExcel = strExcel;
		}

		public bool ReadExcel()
		{
			if (null == m_excelReader)
			{
				m_excelReader = new ExcelReader(m_strFullExcelName);
			}
			// 
			int nSheets = m_excelReader.GetSheetsCount();
			for (int i = 0; i < nSheets; i++)
			{
				string strBoxID = m_excelReader.GetSheetNameByIndex(i);
				List<BoxInfo> listInfo = m_excelReader.GetSheetData<BoxInfo>(strBoxID);
				//
				cBoxInfo info = new cBoxInfo(strBoxID);
				info.CopyAllBox(listInfo);
				m_listBoxInfo.Add(info);
			}

			return true;
		}

		public bool CheckData()
		{
			bool bReturn = true;
			foreach (cBoxInfo info in m_listBoxInfo)
			{
				if (!info.CheckData())
				{
					bReturn = false;
				}
			}

			return bReturn;
		}

		public bool SaveToLua()
		{
			// Create the lua director 
			string targetDir = m_strDirector + "\\..\\lua\\box";
			if (!Directory.Exists(targetDir))
			{
				Directory.CreateDirectory(targetDir);
			}

			// UTF8 NO BOM 
			foreach (cBoxInfo info in m_listBoxInfo)
			{
				// lua file 
				string strLua = targetDir + "\\" + CustomDefine.BOX_LUA_NAME_PREFIX + info.BoxID + ".lua";
				m_sWriter = new StreamWriter(strLua, false, CustomDefine.WRITE_ENCODING);
				if (null != m_sWriter)
				{
					info.SaveData(m_sWriter);
					m_sWriter.Flush();
					m_sWriter.Close();
				}			
			}

			return true;
		}

		class BoxInfo
		{
			[ExcelHeader("男奖励物品编号")]
			public int malereward_itemid { get; set; }
			[ExcelHeader("男奖励物品名称")]
			public string malereward_itemname { get; set; }
			[ExcelHeader("男奖励物品数量")]
			public int malereward_itemcount { get; set; }
			[ExcelHeader("男奖励物品时限")]
			public int malereward_itemvalidity { get; set; }
			[ExcelHeader("男奖励权重")]
			public float malereward_weight { get; set; }
			[ExcelHeader("女奖励物品编号")]
			public int femalereward_itemid { get; set; }
			[ExcelHeader("女奖励物品名称")]
			public string femalereward_itemname { get; set; }
			[ExcelHeader("女奖励物品数量")]
			public int femalereward_itemcount { get; set; }
			[ExcelHeader("女奖励物品时限")]
			public int femalereward_itemvalidity { get; set; }
			[ExcelHeader("女奖励权重")]
			public float femalereward_weight { get; set; }
			[ExcelHeader("公告")]
			public byte announce { get; set; }
			
		}

		class cBoxInfo
		{
			string m_strID = "";
			List<BoxInfo> m_listReward = null;

			public cBoxInfo(string strID)
			{
				m_strID = strID;
			}

			public string BoxID
			{
				get
				{
					return m_strID;
				}
			}
			
			public void AddOneBox(BoxInfo info)
			{
				if (m_listReward != null)
				{
					m_listReward.Add(info);
				}
			}

			public void CopyAllBox(List<BoxInfo> listInfo)
			{
				m_listReward = new List<BoxInfo>(listInfo);
			}

			public void SaveData(StreamWriter sw)
			{
				int nWeight = 10000;	// 权重（硬编码）
				int nGroup = 1;			// 群组ID（硬编码）
				// write to lua 
				sw.WriteLine();
				sw.WriteLine("--物品组概率,万分比");
				string strTemp = "local BoxMaleGroup_" + BoxID + " =";
				sw.WriteLine(strTemp);
				sw.WriteLine("{");
				sw.WriteLine("\t[1] = " + nWeight.ToString() + ",");
				sw.WriteLine("};");
				strTemp = "local BoxFemaleGroup_" + BoxID + " =";
				sw.WriteLine(strTemp);
				sw.WriteLine("{");
				sw.WriteLine("\t[1] = " + nWeight.ToString() + ",");
				sw.WriteLine("};");
				//////////////////////////////////////////////////
				sw.WriteLine("--组物品概率,权重");
				sw.WriteLine("--announce=true/false");
				string strMaleReward = "local BoxMaleItem_" + BoxID + " = " + Environment.NewLine + "{" + Environment.NewLine;
				string strFemaleReward = "local BoxFemaleItem_" + BoxID + " = " + Environment.NewLine + "{" + Environment.NewLine;
				string strMaleItem = string.Empty;
				string strFemaleItem = string.Empty;
				int nItemCount = 0;
				foreach (BoxInfo info in m_listReward)
				{
					strMaleItem = "\t[" + (nItemCount + 1).ToString() + "] = { group = " + nGroup.ToString() 
						+ ", item = \"" + info.malereward_itemid.ToString()
						+ CustomDefine.Separator_In_Item + info.malereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.malereward_itemvalidity).ToString()
						+ "\", announce = " + CustomFunc.ConvertBoolen2String(info.announce != 0)
						+ ", rate = " + (info.malereward_weight * nWeight).ToString()
						+ "}," + Environment.NewLine;
					strFemaleItem = "\t[" + (nItemCount + 1).ToString() + "] = { group = " + nGroup.ToString()
						+ ", item = \"" + info.femalereward_itemid.ToString()
						+ CustomDefine.Separator_In_Item + info.femalereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.femalereward_itemvalidity).ToString()
						+ "\", announce = " + CustomFunc.ConvertBoolen2String(info.announce != 0)
						+ ", rate = " + (info.femalereward_weight * nWeight).ToString()
						+ "}," + Environment.NewLine;
					nItemCount++;
					strMaleReward += strMaleItem;
					strFemaleReward += strFemaleItem;
				}
				strMaleReward += "};";
				strFemaleReward += "};";
				sw.WriteLine(strMaleReward);
				sw.WriteLine();
				sw.WriteLine(strFemaleReward);
				sw.WriteLine();
				//////////////////////////////////////////////////
				sw.WriteLine("function AddBoxMaleGroup(index, value)");
				sw.WriteLine("\tif value ~= nil then");
				strTemp = "\t\tAddBoxGroupInfo( " + BoxID.ToString() + ", index, value, true );";
				sw.WriteLine(strTemp);
				sw.WriteLine("\tend");
				sw.WriteLine("end");
				sw.WriteLine();

				sw.WriteLine("function AddBoxFemaleGroup(index, value)");
				sw.WriteLine("\tif value ~= nil then");
				strTemp = "\t\tAddBoxGroupInfo( " + BoxID.ToString() + ", index, value, false );";
				sw.WriteLine(strTemp);
				sw.WriteLine("\tend");
				sw.WriteLine("end");
				sw.WriteLine();
				//////////////////////////////////////////////////
				sw.WriteLine("function AddBoxMaleItem(index, value)");
				sw.WriteLine("\tif value ~= nil then");
				sw.WriteLine("\t\tlocal groupid = value[\"group\"];");
				sw.WriteLine("\t\tlocal iteminfo = value[\"item\"];");
				sw.WriteLine("\t\tlocal announce = value[\"announce\"];");
				sw.WriteLine("\t\tlocal itemrate = value[\"rate\"];");
				sw.WriteLine();
				strTemp = "\t\tAddBoxItemInfo( " + BoxID.ToString() + ", groupid, iteminfo, announce, itemrate, true );";
				sw.WriteLine(strTemp);
				sw.WriteLine("\tend");
				sw.WriteLine("end");
				sw.WriteLine();

				sw.WriteLine("function AddBoxFemaleItem(index, value)");
				sw.WriteLine("\tif value ~= nil then");
				sw.WriteLine("\t\tlocal groupid = value[\"group\"];");
				sw.WriteLine("\t\tlocal iteminfo = value[\"item\"];");
				sw.WriteLine("\t\tlocal announce = value[\"announce\"];");
				sw.WriteLine("\t\tlocal itemrate = value[\"rate\"];");
				sw.WriteLine();
				strTemp = "\t\tAddBoxItemInfo( " + BoxID.ToString() + ", groupid, iteminfo, announce, itemrate, false );";
				sw.WriteLine(strTemp);
				sw.WriteLine("\tend");
				sw.WriteLine("end");
				sw.WriteLine();

				//////////////////////////////////////////////////
				strTemp = "table.foreach( BoxMaleGroup_" + BoxID.ToString() + ", AddBoxMaleGroup );";
				sw.WriteLine(strTemp);
				strTemp = "table.foreach( BoxFemaleGroup_" + BoxID.ToString() + ", AddBoxFemaleGroup );";
				sw.WriteLine(strTemp);
				sw.WriteLine();

				strTemp = "table.foreach( BoxMaleItem_" + BoxID.ToString() + ", AddBoxMaleItem );";
				sw.WriteLine(strTemp);
				strTemp = "table.foreach( BoxFemaleItem_" + BoxID.ToString() + ", AddBoxFemaleItem );";
				sw.WriteLine(strTemp);
				sw.WriteLine();
			}

			public bool CheckData()
			{
				bool bReturn = true;
				foreach (BoxInfo info in m_listReward)
				{
					bool bCheckMale = ItemMgr.Instance.CheckItemAndLogError((ushort)info.malereward_itemid, CustomDefine.BOX_EXCEL_NAME);
					bool bCheckFemale = ItemMgr.Instance.CheckItemAndLogError((ushort)info.femalereward_itemid, CustomDefine.BOX_EXCEL_NAME);
					if (!bCheckFemale || !bCheckMale)
					{
						bReturn = false;
					}
				}
				return bReturn;
			}
		}
	}
}
