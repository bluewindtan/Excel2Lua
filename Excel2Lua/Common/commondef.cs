﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel2Lua
{
	public class CustomDefine
	{
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
		
		// sheet 
		public static string SHEET_BASE = "通用";


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
		public static string Packet_EXCEL_NAME = "礼包";
		public static string Packet_Lua_NAME_Prefix = "Packet_";

		//////////////////////////////////////////////////////////////
		// 工具文本 
		public static string[] TOOL_TEXT = { "活动脚本转换"
											   , "礼包脚本转换"
								  };
	}

	public enum LuaType
	{
		Activity = 0,	// 活动 
		Packet,			// 礼包
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
				+ m_nValidity.ToString();
		}

	}

}
