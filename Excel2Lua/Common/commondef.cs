using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel2Lua
{
	public class CustomDefine
	{
		// flag of replace item 
		public static string Flag_Replace = "TanFeng";
		// separator of item
		public static string Separator_Item = " | ";
		// prefix of namespace 
		public static string Prefix_NameSpace = "Excel2Lua.";
		//// prefix of E2L base 
		//public static string Prefix_E2LBase = "E2L_";
		// prefix of activity base 
		public static string Prefix_ActivityBase = "AB_";
		
		// sheet 
		public static string SHEET_BASE = "通用";


		//////////////////////////////////////////////////////////////

		public static string[] EXCEL_NAME = { "准点在线"
												, "累计在线"
								  };
		public static string[] LUA_NAME = { "intimeonline"
											  , "OnlineReward"
								  };
	}

	//////////////////////////////////////////////////
	public class Lua_Item
	{
		public int m_nID { get; set; }
		public int m_nCount { get; set; }
		public int m_nValidity { get; set; }

		public Lua_Item(int nID, int nCount, int nValidity)
		{
			m_nID = nID;
			m_nCount = nCount;
			m_nValidity = nValidity;
		}

		public string BuildLua()
		{
			return m_nID.ToString() + "," + m_nCount.ToString() + "," + m_nValidity.ToString();
		}

	}

}
