using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;
using System.IO;

namespace Excel2Lua
{
	public class AB_ExchangeItem : ActivityBase
	{
		List<BaseSheetInfo> m_listBase = null;
		List<ExchangeItemInfo> m_listOwn = null;

		public AB_ExchangeItem(string strExcelSheet)
			: base(strExcelSheet)
		{

		}

		public override bool ReadData(ExcelReader reader)
		{			
			if (! base.ReadData(reader))
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
			m_listOwn = reader.GetSheetData<ExchangeItemInfo>(m_strExcelSheet);

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

			sw.WriteLine("--活动显示起始时间 ");
			strTemp = "local beginExhibitTime = \"" + bsInfo.beginExhibitTime + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local endExhibitTime = \"" + bsInfo.endExhibitTime + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动起始时间 ");
			strTemp = "local exchangeBeginTime = \"" + bsInfo.exchangeBeginTime + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local exchangeEndTime = \"" + bsInfo.exchangeEndTime + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动名称 ");
			strTemp = "local activity_Title = \"" + bsInfo.activity_Title + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();
			sw.WriteLine("--活动内容 ");
			strTemp = "local exchangeActivityIntro = \"" + bsInfo.exchangeActivityIntro + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动起始公告 ");
			strTemp = "local exchangeStartAnnouce = \"" + bsInfo.exchangeStartAnnouce + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local exchangeEndAnnouce = \"" + bsInfo.exchangeEndAnnouce + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--兑换物品 ");
			strTemp = "local exchangeItemType = " + bsInfo.exchangeItemType.ToString() + ";";
			sw.WriteLine(strTemp);
			sw.WriteLine();
			/////////////////////////////////////////////////
			sw.WriteLine("local ExhangeRewardTable = ");
			sw.WriteLine("{");
			/////////////////////////////////////////////////
			int nCount = 1;
			strTemp = "";
			// Loop the data list 
			ExchangeItemInfo infoFirst = null;
			string strMale = string.Empty;
			string strFemale = string.Empty;
			for (int i = 0; i < m_listOwn.Count; i++)
			{
				ExchangeItemInfo info = m_listOwn[i];
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
						sw.WriteLine( infoFirst.BuildLua(nCount, strMale, strFemale) );
						nCount++;
					}
					infoFirst = info;	
					strMale = sMaleItem;
					strFemale = sFemaleItem;
				}
			}
			sw.WriteLine( infoFirst.BuildLua(nCount, strMale, strFemale) );
			/////////////////////////////////////////////////
			sw.WriteLine("}");
			sw.WriteLine();

			strTemp = "function AddExchangeItemTableInfo(index, value)" + Environment.NewLine 
			+ "\tif value ~= nil then" + Environment.NewLine
			+ "\t\tlocal requireItemNum = value[\"requireItemNum\"];" + Environment.NewLine
			+ "\t\tlocal maleReward = value[\"maleReward\"];" + Environment.NewLine
			+ "\t\tlocal femaleReward = value[\"femaleReward\"];" + Environment.NewLine
			+ "\t\tlocal money = value[\"money\"];" + Environment.NewLine
			+ "\t\tAddExchangeItemReward(index, requireItemNum, maleReward, femaleReward, money);" + Environment.NewLine
			+ "\tend" + Environment.NewLine 
			+ "end" + Environment.NewLine 
			+ Environment.NewLine
			+ "function AddExchangeItemInfo(weight)" + Environment.NewLine
			+ "\tif weight ~= nil then" + Environment.NewLine
			+ "\t\tAddExchangeItemBriefInfo(exhibit, weight,regularImageName, thumbnailName, exchangeItemType, beginExhibitTime, endExhibitTime, exchangeBeginTime, exchangeEndTime, activity_Title, exchangeActivityIntro, exchangeStartAnnouce, exchangeEndAnnouce);" + Environment.NewLine
			+ "\t\ttable.foreach(ExhangeRewardTable, AddExchangeItemTableInfo);" + Environment.NewLine
			+ "\tend" + Environment.NewLine 
			+ "end" + Environment.NewLine;
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
			public string activity_Title { get; set; }
			[ExcelHeader("活动内容")]
			public string exchangeActivityIntro { get; set; }
			[ExcelHeader("显示开始时间")]
			public string beginExhibitTime { get; set; }
			[ExcelHeader("显示结束时间")]
			public string endExhibitTime { get; set; }
			[ExcelHeader("活动开始时间")]
			public string exchangeBeginTime { get; set; }
			[ExcelHeader("活动结束时间")]
			public string exchangeEndTime { get; set; }
			[ExcelHeader("开始公告")]
			public string exchangeStartAnnouce { get; set; }
			[ExcelHeader("结束公告")]
			public string exchangeEndAnnouce { get; set; }
			[ExcelHeader("兑换物品")]
			public int exchangeItemType { get; set; }

		}

		class ExchangeItemInfo
		{
			[ExcelHeader("需求兑换物品的数量")]
			public int requireItemNum { get; set; }
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

			public bool IsSameKind(ExchangeItemInfo info)
			{
				if (null == info)
				{
					return false;
				}

				return requireItemNum == info.requireItemNum;
			}
			
			public string BuildLua(int nCount, string strMale, string strFemale)
			{

				return "\t[" + nCount.ToString() + "] = { requireItemNum = " + requireItemNum
					+ ", maleReward = \"" + strMale + "\", femaleReward = \"" + strFemale + "\", money = " + money
					+ " },";
			}

		}
	}
}
