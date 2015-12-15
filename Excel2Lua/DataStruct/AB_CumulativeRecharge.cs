using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;
using System.IO;
using System.Collections;

namespace Excel2Lua
{
	public class AB_CumulativeRecharge : ActivityBase
	{
		List<BaseSheetInfo> m_listBase = null;
		List<CumulativeRechargeInfo> m_listOwn = null;

		public AB_CumulativeRecharge(string strExcelSheet)
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
			m_listOwn = reader.GetSheetData<CumulativeRechargeInfo>(m_strExcelSheet);

			return true;
		}

		public override bool CheckData()
		{
			bool bReturn = true;
			foreach (CumulativeRechargeInfo info in m_listOwn)
			{
				bool bCheckMale = ItemMgr.Instance.CheckItemAndLogError(m_strExcelSheet, info.malereward_itemid, info.malereward_itemcount, info.malereward_itemvalidity);
				bool bCheckFemale = ItemMgr.Instance.CheckItemAndLogError(m_strExcelSheet, info.femalereward_itemid, info.femalereward_itemcount, info.femalereward_itemvalidity);
				if (!bCheckFemale || !bCheckMale)
				{
					bReturn = false;
				}
			}

			return bReturn;
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
			sw.WriteLine("--活动内容 ");
			strTemp = "local activity_content = \"" + bsInfo.activity_content + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("-- 活动起始时间 ");
			strTemp = "local activity_begin_time = \"" + bsInfo.activity_begin_time + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local activity_end_time  = \"" + bsInfo.activity_end_time + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("-- 累计充值起始时间 ");
			strTemp = "local recharge_begin_time = \"" + bsInfo.recharge_begin_time + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local recharge_end_time = \"" + bsInfo.recharge_end_time + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动起始公告 ");
			strTemp = "local recharge_begin_annouce = \"" + bsInfo.recharge_begin_annouce + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local recharge_end_annouce = \"" + bsInfo.recharge_end_annouce + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			/////////////////////////////////////////////////
			sw.WriteLine("local CummulativeRechargeInfo = ");
			sw.WriteLine("{");
			/////////////////////////////////////////////////
			int nCount = 1;
			strTemp = "";
			// Loop the data list 
			CumulativeRechargeInfo infoFirst = null;
			string strMale = string.Empty;
			string strFemale = string.Empty;
			for (int i = 0; i < m_listOwn.Count; i++)
			{
				CumulativeRechargeInfo info = m_listOwn[i];
				// male item 
				string sMaleItem = string.Empty;
				bool bMaleValid = info.malereward_itemid > 0;
				if (bMaleValid)
				{
					sMaleItem = info.malereward_itemid.ToString()
						+ CustomDefine.Separator_In_Item + info.malereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.malereward_itemvalidity).ToString();
				}
				// female item 
				string sFemaleItem = string.Empty;
				bool bFemaleValid = info.femalereward_itemid > 0;
				if (bFemaleValid)
				{
					sFemaleItem = info.femalereward_itemid.ToString()
						+ CustomDefine.Separator_In_Item + info.femalereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.femalereward_itemvalidity).ToString();
				}
				// check the same kind 
				if (info.IsSameKind(infoFirst))
				{
					if (bMaleValid)
					{
						strMale += CustomDefine.Separator_Between_Item + sMaleItem;
					}
					if (bFemaleValid)
					{
						strFemale += CustomDefine.Separator_Between_Item + sFemaleItem;
					}
				}
				else
				{
					if (strMale != string.Empty)
					{
						sw.WriteLine(infoFirst.BuildLua(nCount, strMale, strFemale));
						nCount++;
					}
					infoFirst = info;
					strMale = sMaleItem;
					strFemale = sFemaleItem;
				}
			}
			sw.WriteLine(infoFirst.BuildLua(nCount, strMale, strFemale));
			/////////////////////////////////////////////////
			sw.WriteLine("}");
			sw.WriteLine();

			strTemp = "function AddCummulativeRechargeTableInfo(index, value)" + Environment.NewLine
			+ "\tif value ~= nil then" + Environment.NewLine
			+ "\t\tlocal requirenum = value[\"requirenum\"];" + Environment.NewLine
			+ "\t\tlocal malereward = value[\"malereward\"];" + Environment.NewLine
			+ "\t\tlocal femalereward = value[\"femalereward\"];" + Environment.NewLine
			+ "\t\tlocal money = value[\"money\"];" + Environment.NewLine
			+ "\t\tlocal bindmcoin = value[\"bindmcoin\"];" + Environment.NewLine
			+ Environment.NewLine
			+ "\t\tAddCummulativeRecharge(index, requirenum, malereward, femalereward, money, bindmcoin );" + Environment.NewLine
			+ "\tend" + Environment.NewLine
			+ "end" + Environment.NewLine
			+ Environment.NewLine
			+ "function AddCumulativeInfo(weight)" + Environment.NewLine
			+ "\tif weight ~= nil then" + Environment.NewLine
			+ "\t\tAddCumulativeBrief(exhibit, weight, regularImageName, thumbnailName, activity_title, activity_content, activity_begin_time, activity_end_time, recharge_begin_time, recharge_end_time, recharge_begin_annouce, recharge_end_annouce);" + Environment.NewLine
			+ "\t\ttable.foreach(CummulativeRechargeInfo, AddCummulativeRechargeTableInfo);" + Environment.NewLine
			+ "\tend" + Environment.NewLine
			+ "end" + Environment.NewLine
			+ Environment.NewLine;
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
			[ExcelHeader("开始时间")]
			public string activity_begin_time { get; set; }
			[ExcelHeader("结束时间")]
			public string activity_end_time { get; set; }
			[ExcelHeader("累计充值开始时间")]
			public string recharge_begin_time { get; set; }
			[ExcelHeader("累计充值结束时间")]
			public string recharge_end_time { get; set; }
			[ExcelHeader("开始公告")]
			public string recharge_begin_annouce { get; set; }
			[ExcelHeader("结束公告")]
			public string recharge_end_annouce { get; set; }
		}

		class CumulativeRechargeInfo
		{
			[ExcelHeader("累计数量")]
			public string requirenum { get; set; }
			[ExcelHeader("男奖励物品编号")]
			public uint malereward_itemid { get; set; }
			[ExcelHeader("男奖励物品名称")]
			public string malereward_itemname { get; set; }
			[ExcelHeader("男奖励物品数量")]
			public int malereward_itemcount { get; set; }
			[ExcelHeader("男奖励物品时限")]
			public int malereward_itemvalidity { get; set; }
			[ExcelHeader("女奖励物品编号")]
			public uint femalereward_itemid { get; set; }
			[ExcelHeader("女奖励物品名称")]
			public string femalereward_itemname { get; set; }
			[ExcelHeader("女奖励物品数量")]
			public int femalereward_itemcount { get; set; }
			[ExcelHeader("女奖励物品时限")]
			public int femalereward_itemvalidity { get; set; }
			[ExcelHeader("金券奖励")]
			public uint money { get; set; }
			[ExcelHeader("绑定M币奖励")]
			public uint bindmcoin { get; set; }

			public bool IsSameKind(CumulativeRechargeInfo info)
			{
				if (null == info)
				{
					return false;
				}
				if (requirenum != info.requirenum)
				{
					return false;
				}

				return true;
			}

			public string BuildLua(int nCount, string strMale, string strFemale)
			{

				return "\t[" + nCount.ToString() + "] = { requirenum = " + requirenum
					+ ", malereward = \"" + strMale + "\", femalereward = \"" + strFemale + "\", money = " + money
					+ ", bindmcoin = " + bindmcoin
					+ " },";
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

			public void AddItem(uint nMaleItemID, int nMaleItemCount, int nMaleItemValidity
				, uint nFemaleItemID, int nFemaleItemCount, int nFemaleItemValidity)
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
