using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;
using System.IO;
using System.Collections;

namespace Excel2Lua
{
	public class AB_FresherActivity : ActivityBase
	{
		List<BaseSheetInfo> m_listBase = null;
		List<FresherActivityInfo> m_listOwn = null;

		public AB_FresherActivity(string strExcelSheet)
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
			m_listOwn = reader.GetSheetData<FresherActivityInfo>(m_strExcelSheet);

			return true;
		}

		public override bool CheckData()
		{
			bool bReturn = true;
			foreach (FresherActivityInfo info in m_listOwn)
			{
				bool bCheckMale = ItemMgr.Instance.CheckItemAndLogError((ushort)info.malereward_itemid, m_strExcelSheet);
				bool bCheckFemale = ItemMgr.Instance.CheckItemAndLogError((ushort)info.femalereward_itemid, m_strExcelSheet);
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
			sw.WriteLine("--活动名称和内容 ");
			string strTemp = "local activityName = \"" + bsInfo.activityName + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local activityContent = \"" + bsInfo.activityContent + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--背包已满时，通过邮件发送奖励，邮件标题和内容 ");
			strTemp = "local mailTitle = \"" + bsInfo.mailTitle + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local mailContent = \"" + bsInfo.mailContent + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动时长, 单位：天 ");
			strTemp = "local activityTime = " + bsInfo.activityTime + ";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			/////////////////////////////////////////////////
			sw.WriteLine("--N天内的奖励配置，第几天及对应的奖励 ");
			sw.WriteLine("local activityReward = ");
			sw.WriteLine("{");
			/////////////////////////////////////////////////
			int nCount = 1;
			strTemp = "";
			// Loop the data list 
			FresherActivityInfo infoFirst = null;
			string strMale = string.Empty;
			string strFemale = string.Empty;
			string strVIPMale = string.Empty;
			string strVIPFemale = string.Empty;
			string strCumulationLua = string.Empty;
			for (int i = 0; i < m_listOwn.Count; i++)
			{
				FresherActivityInfo info = m_listOwn[i];
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
				// VIP male item 
				string sVIPMaleItem = string.Empty;
				bool bVIPMaleValid = info.vipmalereward_itemid > 0;
				if (bVIPMaleValid)
				{
					sVIPMaleItem = info.vipmalereward_itemid.ToString()
						+ CustomDefine.Separator_In_Item + info.vipmalereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.vipmalereward_itemvalidity).ToString();
				}
				// VIP female item 
				string sVIPFemaleItem = string.Empty;
				bool bVIPFemaleValid = info.vipfemalereward_itemid > 0;
				if (bVIPFemaleValid)
				{
					sVIPFemaleItem = info.vipfemalereward_itemid.ToString()
						+ CustomDefine.Separator_In_Item + info.vipfemalereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.vipfemalereward_itemvalidity).ToString();
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
					if (bVIPMaleValid)
					{
						strVIPMale += CustomDefine.Separator_Between_Item + sVIPMaleItem;
					}
					if (bVIPFemaleValid)
					{
						strVIPFemale += CustomDefine.Separator_Between_Item + sVIPFemaleItem;
					}
				}
				else
				{
					if (strMale != string.Empty)
					{
						sw.WriteLine(infoFirst.BuildLua(nCount, strMale, strFemale, strVIPMale, strVIPFemale));
						nCount++;
					}
					infoFirst = info;
					strMale = sMaleItem;
					strFemale = sFemaleItem;
					strVIPMale = sVIPMaleItem;
					strVIPFemale = sVIPFemaleItem;
				}
			}
			if (nCount > 7)
			{
				strCumulationLua = infoFirst.BuildLua(strMale, strFemale, strVIPMale, strVIPFemale);
			}
			else
			{
				sw.WriteLine(infoFirst.BuildLua(nCount, strMale, strFemale, strVIPMale, strVIPFemale));
			}
			/////////////////////////////////////////////////
			sw.WriteLine("};");
			sw.WriteLine();

			sw.WriteLine("--补签需要的金券 ");
			strTemp = "local recvAgainMoney = " + bsInfo.recvAgainMoney + ";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--在线多少天可以获得全签奖励, 单位：天 ");
			strTemp = "local cumulationTime = " + bsInfo.cumulationTime + ";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			/////////////////////////////////////////////////
			sw.WriteLine("--全签奖励配置 ");
			strTemp = "local cumulationReward = " + strCumulationLua;
			sw.WriteLine(strTemp);
			/////////////////////////////////////////////////

			strTemp = "AddFresherActivity(activityName, activityContent, mailTitle, mailContent, recvAgainMoney, activityTime, cumulationTime);" + Environment.NewLine
			+ "AddFresherActivityCumulationReward(cumulationReward.maleItem, cumulationReward.femaleItem, cumulationReward.money, cumulationReward.bindCoin," + Environment.NewLine
			+ "\t\tcumulationReward.vipMaleItem, cumulationReward.vipFemaleItem, cumulationReward.vipMoney, cumulationReward.vipBindCoin);" + Environment.NewLine
			+ Environment.NewLine
			+ "function AddFresherActivityRewardRoutine(index, value)" + Environment.NewLine
			+ "\tif value ~= nil then" + Environment.NewLine
			+ "\t\tAddFresherActivityReward(value.day, value.maleItem, value.femaleItem, value.money, value.bindCoin," + Environment.NewLine
			+ "\t\t\tvalue.vipMaleItem,value.vipFemaleItem,value.vipMoney,value.vipBindCoin);" + Environment.NewLine
			+ "\tend" + Environment.NewLine
			+ "end" + Environment.NewLine
			+ Environment.NewLine
			+ "table.foreach(activityReward, AddFresherActivityRewardRoutine);" + Environment.NewLine
			+ Environment.NewLine;
			sw.WriteLine(strTemp);

			sw.Flush();
			return true;
		}

		//////////////////////////////////////////////////////////////////////////////////////////////////
		class BaseSheetInfo
		{
			[ExcelHeader("活动名称")]
			public string activityName { get; set; }
			[ExcelHeader("活动内容")]
			public string activityContent { get; set; }
			[ExcelHeader("奖励邮件标题")]
			public string mailTitle { get; set; }
			[ExcelHeader("奖励邮件内容")]
			public string mailContent { get; set; }
			[ExcelHeader("活动时长")]
			public int activityTime { get; set; }
			[ExcelHeader("补签所需金券")]
			public int recvAgainMoney { get; set; }
			[ExcelHeader("全签奖励所需天数")]
			public int cumulationTime { get; set; }
		}

		class FresherActivityInfo
		{
			[ExcelHeader("天数")]
			public int day { get; set; }
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
			public int money { get; set; }
			[ExcelHeader("绑定M币奖励")]
			public int bindmcoin { get; set; }
			[ExcelHeader("VIP男奖励物品编号")]
			public int vipmalereward_itemid { get; set; }
			[ExcelHeader("VIP男奖励物品名称")]
			public string vipmalereward_itemname { get; set; }
			[ExcelHeader("VIP男奖励物品数量")]
			public int vipmalereward_itemcount { get; set; }
			[ExcelHeader("VIP男奖励物品时限")]
			public int vipmalereward_itemvalidity { get; set; }
			[ExcelHeader("VIP女奖励物品编号")]
			public int vipfemalereward_itemid { get; set; }
			[ExcelHeader("VIP女奖励物品名称")]
			public string vipfemalereward_itemname { get; set; }
			[ExcelHeader("VIP女奖励物品数量")]
			public int vipfemalereward_itemcount { get; set; }
			[ExcelHeader("VIP女奖励物品时限")]
			public int vipfemalereward_itemvalidity { get; set; }
			[ExcelHeader("VIP金券奖励")]
			public int vipmoney { get; set; }
			[ExcelHeader("VIP绑定M币奖励")]
			public int vipbindmcoin { get; set; }

			public bool IsSameKind(FresherActivityInfo info)
			{
				if (null == info)
				{
					return false;
				}
				if (day != info.day)
				{
					return false;
				}

				return true;
			}

			public string BuildLua(int nCount, string strMale, string strFemale, string strVIPMale, string strVIPFemale)
			{

				return "\t[" + nCount.ToString() + "] = { day = " + day.ToString()
					+ ", maleItem = \"" + strMale + "\", femaleItem = \"" + strFemale + "\", money = " + money
					+ ", bindCoin = " + bindmcoin + ","
					+ Environment.NewLine
					+ "\t\t\tvipMaleItem = \"" + strVIPMale + "\", vipFemaleItem = \"" + strVIPFemale + "\", vipMoney = " + vipmoney
					+ ", vipBindCoin = " + vipbindmcoin
					+ " };";
			}

			public string BuildLua(string strMale, string strFemale, string strVIPMale, string strVIPFemale)
			{

				return "{ maleItem = \"" + strMale + "\", femaleItem = \"" + strFemale + "\", money = " + money
					+ ", bindCoin = " + bindmcoin + ","
					+ Environment.NewLine
					+ "\t\t\tvipMaleItem = \"" + strVIPMale + "\", vipFemaleItem = \"" + strVIPFemale + "\", vipMoney = " + vipmoney
					+ ", vipBindCoin = " + vipbindmcoin
					+ " };";
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
