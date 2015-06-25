using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;
using System.IO;

namespace Excel2Lua
{
	public class ActivityBase
	{
		protected string m_strExcelSheet = "";

		protected ActivityBase(string strExcelSheet)
		{
			m_strExcelSheet = strExcelSheet;
		}

		public virtual bool ReadData(ExcelReader reader)
		{
			if (null == reader)
			{
				return false;
			}
			
			return true;
		}

        public virtual bool SaveData(StreamWriter sw)
        {
			if (null == sw)
			{
				return false;
			}
			return true;
        }
	}

}
