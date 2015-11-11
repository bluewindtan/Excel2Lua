using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Excel2Lua
{
	public class E2LData
	{
		public E2LData(string e, string l)
		{
			strExcel = e;
			strLua = l;
		}
		public string strExcel;
		public string strLua;
	}

	public sealed class E2LDataMgr
	{
		static E2LDataMgr s_instance = null;
		static readonly object lockTemp = new object();

		public static E2LDataMgr Instance
		{
			get
			{
				lock (lockTemp)
				{
					if (null == s_instance)
					{
						s_instance = new E2LDataMgr();
					}
					return s_instance;
				}
			}
		}

		public E2LBase CreateOperater(string m_strFullExcelName)
		{
			E2LBase operE2L = null;

			FileInfo fi = new FileInfo(m_strFullExcelName);
			if (!fi.Exists)
			{
				return null;
			}
			string strExcel = fi.Name.Replace(fi.Extension, string.Empty);

			if (m_dicE2LData.ContainsKey(strExcel))
			{
				string strLua = m_dicE2LData[strExcel].strLua;
				// need to add the Excel2Lua namespace 
				//string strTypeFull = CustomDefine.Prefix_NameSpace + CustomDefine.Prefix_E2LBase + strLua;
				//Type tpE2L = Type.GetType(strTypeFull);
				//operE2L = (E2LBase)Activator.CreateInstance(tpE2L, new string[] { fi.DirectoryName, strExcel, fi.Extension, strLua });
				operE2L = new E2LBase(fi.DirectoryName, strExcel, fi.Extension, strLua);

				string strTypeFull = CustomDefine.Prefix_NameSpace + CustomDefine.Prefix_ActivityBase + strLua;
				Type tpAB = Type.GetType(strTypeFull);
				operE2L.actiBase = (ActivityBase)Activator.CreateInstance(tpAB, new string[] { strExcel });
			}
			else
			{
				throw new Exception(strExcel + " exists, but there is no corresponding operator.");
			}

			return operE2L;
		}

		public bool ConvertE2L(string sDirectory, string sFile)
		{
			// analyze the name of excel file 
			if (sFile.Length <= 0)
			{
				return false;
			}
			string strFullExcelName = sDirectory + "\\" + sFile;
			bool nReturn = false;
			LuaType nLuaType = CustomFunc.JudgeType(sFile);
			nReturn = ConvertE2LByType(sDirectory, sFile, nLuaType);
			return nReturn;
		}

		public bool ConvertE2LByType(string sDirectory, string sFile, LuaType nType)
		{
			// analyze the name of excel file 
			if (sFile.Length <= 0)
			{
				return false;
			}
			string strFullExcelName = sDirectory + "\\" + sFile;
			bool nReturn = false;
			LuaType nLuaType = CustomFunc.JudgeType(sFile);
			if (LuaType.Activity == nType)
			{
				if (nLuaType == nType)
				{
					nReturn = ConvertE2L_Activity(sDirectory, sFile);
				}
			}
			else if (LuaType.Packet == nType)
			{
				if (nLuaType == nType)
				{
					nReturn = ConvertE2L_Packet(sDirectory, sFile);
				}
			}
			else if (LuaType.Box == nType)
			{
				if (nLuaType == nType)
				{
					nReturn = ConvertE2L_Box(sDirectory, sFile);
				}
			}
			else
			{

			}

			return nReturn;
		}

		private bool ConvertE2L_Box(string sDirectory, string sFile)
		{
			if (LuaType.Box != CustomFunc.JudgeType(sFile))
			{
				return true;
			}
			E2L_Box e2lBox = new E2L_Box(sDirectory, sFile);
			if (null == e2lBox)
			{
				return false;
			}
			// read the excel 
			if (!e2lBox.ReadExcel())
			{
				return false;
			}
			// check data 
			if (!e2lBox.CheckData())
			{
				//return false; // continue to process other files
			}
			// save to lua 
			if (!e2lBox.SaveToLua())
			{
				return false;
			}

			return true;
		}

		private bool ConvertE2L_Packet(string sDirectory, string sFile)
		{
			if (LuaType.Packet != CustomFunc.JudgeType(sFile))
			{
				return true;
			}
			E2L_Packet e2lPacket = new E2L_Packet(sDirectory, sFile);
			if (null == e2lPacket)
			{
				return false;
			}
			// read the excel 
			if (!e2lPacket.ReadExcel())
			{
				return false;
			}
			// check data 
			if (!e2lPacket.CheckData())
			{
				//return false; // continue to process other files
			}
			// save to lua 
			if (!e2lPacket.SaveToLua())
			{
				return false;
			}

			return true;
		}

		private bool ConvertE2L_Activity(string sDirectory, string sFile)
		{
			if (LuaType.Activity != CustomFunc.JudgeType(sFile))
			{
				return true;
			}
			string strFullExcelName = sDirectory + "\\" + sFile;
			// create the operator of this excel 
			E2LBase operE2L = E2LDataMgr.Instance.CreateOperater(strFullExcelName);
			if (null == operE2L)
			{
				return false;
			}

			// read the excel 
			if (!operE2L.ReadExcel())
			{
				return false;
			}
			// check data 
			if (!operE2L.CheckData())
			{
				//return false; // continue to process other files
			}
			// save to lua 
			if (!operE2L.SaveToLua())
			{
				return false;
			}
			
			return true;
		}


		//////////////////////////////////////////////////////////////////

		Dictionary<string, E2LData> m_dicE2LData = new Dictionary<string, E2LData>();

		E2LDataMgr()
		{
			if (CustomDefine.EXCEL_NAME.Length != CustomDefine.LUA_NAME.Length)
			{
				throw new Exception("Must be a one-to-one correspondence between Excel and Lua.");
			}
			for (int i = 0; i < CustomDefine.EXCEL_NAME.Length; i++)
			{
				m_dicE2LData.Add(CustomDefine.EXCEL_NAME[i],
					new E2LData(CustomDefine.EXCEL_NAME[1], CustomDefine.LUA_NAME[i]));
			}

			// read item excel file 
			ItemMgr.Instance.ReadExcelItem(Environment.CurrentDirectory + "\\source\\" + CustomDefine.ITEM_EXCEL_FILE);
		}

	}

}
