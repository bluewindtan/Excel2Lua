using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;
using System.IO;
using System.Collections;

namespace Excel2Lua
{
	public class AB_OnlineReward : ActivityBase
	{
		List<BaseSheetInfo> m_listBase = null;
		List<OnlineRewardInfo> m_listOwn = null;

		public AB_OnlineReward(string strExcelSheet)
			: base(strExcelSheet)
		{

		}

		public override bool ReadData(ExcelReader reader)
		{
			if (!base.ReadData(reader))
			{
				return false;
			}
			////
			m_listBase = reader.GetSheetData<BaseSheetInfo>(CustomDefine.SHEET_BASE);
			if (m_listBase.Count > 1)
			{
				throw new Exception("Failed to read the sheet=" + CustomDefine.SHEET_BASE + " of excel=" + reader.ToString());
			}
			////
			m_listOwn = reader.GetSheetData<OnlineRewardInfo>(m_strExcelSheet);

			return true;
		}

		public override bool SaveData(StreamWriter sw)
		{
			if (!base.SaveData(sw))
			{
				return false;
			}
			BaseSheetInfo bsInfo = m_listBase[0];
			// write to lua 
			sw.WriteLine("--是否显示在主页, 0表示不显示，大于0表示显示 ");
			string strTemp = "local exhibit = " + bsInfo.exhibit.ToString() + ";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动图片名称 ");
			strTemp = "local regularImageName = \"" + bsInfo.regularImageName + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local thumbnailName = \"" + bsInfo.thumbnailName + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动名称 ");
			strTemp = "local activity_title = \"" + bsInfo.activity_title + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine("--活动内容，其中时间程序会自动填充 ");
			strTemp = "local activity_content = \"" + bsInfo.activity_content + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动显示时间 ");
			strTemp = "local show_begin_time = \"" + bsInfo.show_begin_time + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local show_end_time = \"" + bsInfo.show_end_time + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			//sw.WriteLine("--活动起始时间 ");
			//strTemp = "local activity_start_time = \"" + bsInfo.activity_start_time + "\";";
			//sw.WriteLine(strTemp);
			//strTemp = "local activity_end_time = \"" + bsInfo.activity_end_time + "\";";
			//sw.WriteLine(strTemp);
			//sw.WriteLine();

			sw.WriteLine("--每日计时开始时间 ");
			string[] strArray = bsInfo.resettime.Split(new char[]{':', '：'});
			if (strArray.Length != 3)
			{
				throw new Exception("OnlineReward: resettime error.");
			}
			strTemp = "local resettime = {hour = " + strArray[0]
				+ ", minute = " + strArray[1]
				+ ", second = " + strArray[2] + "};";
			sw.WriteLine(strTemp);
			sw.WriteLine();
			/////////////////////////////////////////////////
			sw.WriteLine("--玩家物品，activity_end_time与下一个activity_start_time不能完全相同");
			sw.WriteLine("local onlineRewardInfo = ");
			sw.WriteLine("{");
			/////////////////////////////////////////////////
			int nOnlineRewardIndex = 1;
			int nOnlineIndex = 1;
			OnlineRewardInfo infoLast = null;
			Lua_OnlineRewardInfo luaAllInfo = null;
			Lua_Online luaReward = null;
			// Loop the data list 
			for (int i = 0; i < m_listOwn.Count; i++)
			{
				OnlineRewardInfo info = m_listOwn[i];
				int nReturn = info.HasSameTime(infoLast);
				if (2 == nReturn) // 同样的奖励时间，并且同样的累计秒数 
				{
					luaReward.AddItem(info.malereward_itemid, info.malereward_itemcount, info.malereward_itemvalidity
						, info.femalereward_itemid, info.femalereward_itemcount, info.femalereward_itemvalidity);
				}
				else if (1 == nReturn) // 同样的奖励时间，不同的累计秒数 
				{
					luaReward = new Lua_Online(nOnlineIndex++, info.onlinetime, info.moneyreward);
					luaReward.AddItem(info.malereward_itemid, info.malereward_itemcount, info.malereward_itemvalidity
						, info.femalereward_itemid, info.femalereward_itemcount, info.femalereward_itemvalidity);
					luaAllInfo.AddLuaReward(luaReward);
				}
				else // 不同的奖励时间，等于新的奖励开始了  
				{
					if (nOnlineRewardIndex > 1)
					{
						sw.WriteLine(luaAllInfo.BuildLua());
					}
					luaAllInfo = new Lua_OnlineRewardInfo(nOnlineRewardIndex++, info.begintime, info.endtime);
					luaReward = new Lua_Online(nOnlineIndex++, info.onlinetime, info.moneyreward);
					luaReward.AddItem(info.malereward_itemid, info.malereward_itemcount, info.malereward_itemvalidity
						, info.femalereward_itemid, info.femalereward_itemcount, info.femalereward_itemvalidity);
					luaAllInfo.AddLuaReward(luaReward);
				}
				infoLast = info;
			}
			sw.WriteLine(luaAllInfo.BuildLua());

			/////////////////////////////////////////////////
			sw.WriteLine("}");
			sw.WriteLine();

			strTemp = "function AddOnlineRewardActivityInfo(weight)" + Environment.NewLine
			+ "\tlocal resetHour = resettime[\"hour\"];" + Environment.NewLine
			+ "\tlocal restMinute = resettime[\"minute\"];" + Environment.NewLine
			+ "\tlocal resetSeconds = resettime[\"second\"];" + Environment.NewLine
			+ Environment.NewLine
			+ "\tAddOnlineTimeRewardBrief(weight, exhibit, show_begin_time, show_end_time, regularImageName, thumbnailName, activity_title, activity_content, resetHour, restMinute, resetSeconds);" + Environment.NewLine
			+ "end" + Environment.NewLine
			+ Environment.NewLine
			+ "function AddOnlineReward(index, value, activityid, activity_start_time, activity_end_time) " + Environment.NewLine
			+ "\tif value ~= nil then" + Environment.NewLine
			+ "\t\tlocal onlinetime = value[\"onlinetime\"];	" + Environment.NewLine
			+ "\t\tlocal malereward = value[\"malereward\"];" + Environment.NewLine
			+ "\t\tlocal femalereward = value[\"femalereward\"];" + Environment.NewLine
			+ "\t\tlocal money = value[\"money\"];" + Environment.NewLine
			+ Environment.NewLine
			+ "\t\tAddOnlineRewardInfo(activityid, activity_start_time, activity_end_time, index, onlinetime, malereward, femalereward, money);" + Environment.NewLine
			+ "\tend" + Environment.NewLine
			+ "end" + Environment.NewLine
			+ Environment.NewLine
			+ "function AddOnlineRewardLoop(index, value)" + Environment.NewLine
			+ "\tlocal detail = value[\"detail\"];" + Environment.NewLine
			+ "\tlocal activity_start_time = value[\"activity_start_time\"];" + Environment.NewLine
			+ "\tlocal activity_end_time = value[\"activity_end_time\"];" + Environment.NewLine
			+ "\tfor i, v in ipairs(detail) do" + Environment.NewLine
			+ "\t\tAddOnlineReward(i, v, index, activity_start_time, activity_end_time);" + Environment.NewLine
			+ "\tend" + Environment.NewLine
			+ "end" + Environment.NewLine
			+ Environment.NewLine
			+ "table.foreach(onlineRewardInfo, AddOnlineRewardLoop);" + Environment.NewLine;
			sw.WriteLine(strTemp);

			sw.Flush();
			return true;
		}

		//////////////////////////////////////////////////////////////////////////////////////////////////
		class BaseSheetInfo
		{
			[ExcelHeader("主页显示")]
			public byte exhibit { get; set; }
			[ExcelHeader("大图")]
			public string regularImageName { get; set; }
			[ExcelHeader("小图")]
			public string thumbnailName { get; set; }
			[ExcelHeader("活动名称")]
			public string activity_title { get; set; }
			[ExcelHeader("活动内容")]
			public string activity_content { get; set; }
			[ExcelHeader("显示开始时间")]
			public string show_begin_time { get; set; }
			[ExcelHeader("显示结束时间")]
			public string show_end_time { get; set; }
			//[ExcelHeader("活动开始时间")]
			//public string activity_start_time { get; set; }
			//[ExcelHeader("活动结束时间")]
			//public string activity_end_time { get; set; }
			[ExcelHeader("每日计时开始时间")]
			public string resettime { get; set; }
		}

		class OnlineRewardInfo
		{
			[ExcelHeader("开始时间")]
			public string begintime { get; set; }
			[ExcelHeader("结束时间")]
			public string endtime { get; set; }
			[ExcelHeader("累计秒数")]
			public string onlinetime { get; set; }
			[ExcelHeader("男奖励物品编号")]
			public int malereward_itemid { get; set; }
			[ExcelHeader("男奖励物品名称")]
			public string malereward_itemname { get; set; }
			[ExcelHeader("男奖励物品数量")]
			public int malereward_itemcount { get; set; }
			[ExcelHeader("男奖励物品时限")]
			public int malereward_itemvalidity { get; set; }
			[ExcelHeader("女奖励物品编号")]
			public int femalereward_itemid { get; set; }
			[ExcelHeader("女奖励物品名称")]
			public string femalereward_itemname { get; set; }
			[ExcelHeader("女奖励物品数量")]
			public int femalereward_itemcount { get; set; }
			[ExcelHeader("女奖励物品时限")]
			public int femalereward_itemvalidity { get; set; }
			[ExcelHeader("金券奖励")]
			public int moneyreward { get; set; }

			public int HasSameTime(OnlineRewardInfo info)
			{
				if (null == info)
				{
					return -1;
				}
				if (0 != begintime.CompareTo(info.begintime))
				{
					return 0;
				}
				if (0 != endtime.CompareTo(info.endtime))
				{
					return 0;
				}
				if (0 != onlinetime.CompareTo(info.onlinetime))
				{
					return 1;
				}

				return 2;
			}

		}

		////////////////////////////////////////////////// 
		class Lua_Online
		{
			int m_nIndex { get; set; }
			string m_strOnlineTime { get; set; }
			int m_nMoney { get; set; }
			public List<Lua_Item> m_listMaleItem;
			public List<Lua_Item> m_listFemaleItem;

			public Lua_Online(int nIndex, string strOnlineTime, int nMoney)
			{
				m_nIndex = nIndex;
				m_strOnlineTime = strOnlineTime;
				m_nMoney = nMoney;
				m_listMaleItem = new List<Lua_Item>();
				m_listFemaleItem = new List<Lua_Item>();
			}

			public void AddItem(bool bMale, Lua_Item item)
			{
				if (bMale)
				{
					m_listMaleItem.Add(item);
				}
				else
				{
					m_listFemaleItem.Add(item);
				}
			}

			public void AddItem(int nMaleItemID, int nMaleItemCount, int nMaleItemValidity
				, int nFemaleItemID, int nFemaleItemCount, int nFemaleItemValidity)
			{
				AddItem(true, new Lua_Item(nMaleItemID, nMaleItemCount, nMaleItemValidity));
				AddItem(false, new Lua_Item(nFemaleItemID, nFemaleItemCount, nFemaleItemValidity));
			}

			public string BuildLua(int nIndex)
			{
				string strTemp = "\t\t[" + nIndex.ToString() + "] = {onlinetime = " + m_strOnlineTime.ToString();
				//////////////////////////////////////////////////
				string strMale = ", malereward = \"";
				for (int i = 0; i < m_listMaleItem.Count; i++)
				{
					if (0 != i)
					{
						strMale += CustomDefine.Separator_Between_Item;
					}
					strMale += m_listMaleItem[i].BuildLua();
				}
				strTemp += strMale + "\"";
				//////////////////////////////////////////////////
				string strFemale = ", femalereward = \"";
				for (int i = 0; i < m_listFemaleItem.Count; i++)
				{
					if (0 != i)
					{
						strFemale += CustomDefine.Separator_Between_Item;
					}
					strFemale += m_listFemaleItem[i].BuildLua();
				}
				strTemp += strFemale + "\"";
				//////////////////////////////////////////////////
				strTemp += ", money = " + m_nMoney.ToString() + "},";
				strTemp += Environment.NewLine;
				return strTemp;
			}
		}

		class Lua_OnlineRewardInfo
		{
			int m_nIndex { get; set; }
			string m_strBeginTime { get; set; }
			string m_strEndTime { get; set; }
			List<Lua_Online> m_listOnline = null;

			public Lua_OnlineRewardInfo(int nIndex, string strBeginTime, string strEndTime)
			{
				m_nIndex = nIndex;
				m_strBeginTime = strBeginTime;
				m_strEndTime = strEndTime;
				m_listOnline = new List<Lua_Online>();
			}

			public void AddLuaReward(Lua_Online item)
			{
				m_listOnline.Add(item);
			}

			public string BuildLua()
			{
				string strTemp = "\t[" + m_nIndex.ToString() + "] = { activity_start_time = \"" + m_strBeginTime
					+ "\", activity_end_time = \"" + m_strEndTime
					+ "\", detail = {";
				strTemp += Environment.NewLine;
				for (int i = 0; i < m_listOnline.Count; i++)
				{
					strTemp += m_listOnline[i].BuildLua(i + 1);
				}

				strTemp += "\t}},";
				return strTemp;
			}

		}
	}
}
