using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using MSExcelReader;

namespace Excel2Lua
{
	public class E2LBase
	{
		protected string m_strFullExcelName = "";
		protected string m_strDirector = "";		
		protected string m_strExcel = "";			 
		protected string m_strExtension = "";
		protected string m_strLua = "";
		protected string m_strLogWhere = "";		

		public ActivityBase actiBase = null;

		protected ExcelReader m_excelReader = null;
		protected StreamWriter m_sWriter = null;

		public E2LBase(string strDirector, string strExcel, string strExtension, string strLua)
		{
			m_strFullExcelName = strDirector + "\\" + strExcel + strExtension;
			m_strDirector = strDirector;		
			m_strExcel = strExcel;
			m_strExtension = strExtension;
			m_strLua = strLua;
			m_strLogWhere = strExcel;
		}

		public bool ReadExcel()
		{
			if (null == m_excelReader)
			{
				m_excelReader = new ExcelReader(m_strFullExcelName);
			}

			return actiBase.ReadData(m_excelReader);
		}

		public bool CheckData()
		{
			return actiBase.CheckData();
		}

		public bool SaveToLua()
		{
			// Create the lua director  
			string targetDir = m_strDirector + "\\..\\lua";
			if (!Directory.Exists(targetDir))
			{
				Directory.CreateDirectory(targetDir);
			}

			// lua file 
			string strLua = targetDir + "\\" + m_strLua + ".lua";

			// UTF8 NO BOM 
			m_sWriter = new StreamWriter(strLua, false, CustomDefine.WRITE_ENCODING);
			if (null == m_sWriter)
			{
				return false;
			}
			if (!actiBase.SaveData(m_sWriter))
			{
				return false;
			}
			m_sWriter.Flush();
			m_sWriter.Close();

			return true;
		}
	}
}
