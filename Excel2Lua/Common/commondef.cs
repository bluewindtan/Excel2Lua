using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Excel2Lua
{
	public class CustomDefine
	{
		//////////////////////////////////////////////////////////////
		// separator between items
		public static string Separator_Between_Item = " | ";
		// separator in item
		public static string Separator_In_Item = ",";
		// prefix of namespace 
		public static string Prefix_NameSpace = "Excel2Lua.";
		//// prefix of E2L base 
		//public static string Prefix_E2LBase = "E2L_";
		// prefix of activity base 
		public static string Prefix_ActivityBase = "AB_";
		// name of log file
		public static string LOG_NAME = "E2LLog";
		// suffix of log file
		public static string SUFFIX_LOG = ".log";
		
		// sheet 
		public static string SHEET_BASE = "通用";

		//////////////////////////////////////////////////////////////
		// path of item excel file
		public static string ITEM_EXCEL_FILE = "物品表.xlsx";

		//////////////////////////////////////////////////////////////
		// encoding
		public static Encoding WRITE_ENCODING = new UTF8Encoding(false);

		//////////////////////////////////////////////////////////////

		public static string[] EXCEL_NAME = { "准点在线"
												, "累计在线"
												, "兑换活动"
												, "累计充值"
												, "累计消费"
												, "七天乐"
								  };
		public static string[] LUA_NAME = { "intimeonline"
											  , "OnlineReward"
											  , "ExchangeItem"
											  , "CumulativeRecharge"
											  , "CumulativeSpend"
											  , "FresherActivity"
								  };

		//////////////////////////////////////////////////////////////
		// 礼包
		public static string PACKET_EXCEL_NAME = "礼包";
		public static string PACKET_LUA_NAME_PREFIX = "Packet_";

		// 宝箱
		public static string BOX_EXCEL_NAME = "宝箱";
		public static string BOX_LUA_NAME_PREFIX = "Box_";

		//////////////////////////////////////////////////////////////
		// 工具文本 
		public static string[] TOOL_TEXT = { "活动脚本转换"
											   , "礼包脚本转换"
											   , "宝箱脚本转换"
								  };
	}

	public class CustomFunc
	{
		// 将时限中的天数转换为秒数 
		public static int Day2Second(int nValidity)
		{
			if (nValidity <= 0)
			{
				return nValidity;
			}

			return nValidity * 24 * 3600;
		}

		public static LuaType JudgeType(string sFile)
		{
			LuaType nType = LuaType.Activity;
			
			// 礼包
			if (sFile.Contains(CustomDefine.PACKET_EXCEL_NAME))
			{
				nType = LuaType.Packet;
			}
			// 宝箱
			if (sFile.Contains(CustomDefine.BOX_EXCEL_NAME))
			{
				nType = LuaType.Box;
			}
			// 物品表 
			if (sFile.Contains(CustomDefine.ITEM_EXCEL_FILE))
			{
				nType = LuaType.Item;
			}

			return nType;
		}

		public static string ConvertBoolen2String(bool bValue)
		{
			return bValue ? "true" : "false";
		}

	}

	public enum LuaType
	{
		Activity = 0,	// 活动 
		Packet,			// 礼包
		Box,			// 宝箱


		Item,			// 物品表 
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
			return m_nID.ToString() + CustomDefine.Separator_In_Item
				+ m_nCount.ToString() + CustomDefine.Separator_In_Item 
				+ CustomFunc.Day2Second(m_nValidity).ToString();
		}

	}

}
