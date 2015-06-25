using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;
using System.IO;

namespace Excel2Lua
{
	public class AB_intimeonline : ActivityBase
	{
		List<BaseSheetInfo> m_listBase = null;
		List<IntimeOnlineInfo> m_listOwn = null;

		public AB_intimeonline(string strExcelSheet)
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
			m_listOwn = reader.GetSheetData<IntimeOnlineInfo>(m_strExcelSheet);

			return true;
		}

		public override bool SaveData(StreamWriter sw)
		{
			if (!base.SaveData(sw))
			{
				return false;
			}
			BaseSheetInfo bsInfo = m_listBase[0];
			bsInfo.CorrectData();
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

			sw.WriteLine("--活动显示起始时间 ");
			strTemp = "local show_time_begin = \"" + bsInfo.show_time_begin + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local show_time_end = \"" + bsInfo.show_time_end + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();

			sw.WriteLine("--活动起始公告 ");
			strTemp = "local activity_begin_announce = \"" + bsInfo.activity_begin_announce + "\";";
			sw.WriteLine(strTemp);
			strTemp = "local activity_end_announce = \"" + bsInfo.activity_end_announce + "\";";
			sw.WriteLine(strTemp);
			sw.WriteLine();
			/////////////////////////////////////////////////
			sw.WriteLine("local InTimeOnlineActivityData = ");
			sw.WriteLine("{");
			/////////////////////////////////////////////////
			int nCount = 1;
			strTemp = "";
			// Loop the data list 
			IntimeOnlineInfo infoFirst = null;
			string strMale = string.Empty;
			string strFemale = string.Empty;
			for (int i = 0; i < m_listOwn.Count; i++)
			{
				IntimeOnlineInfo info = m_listOwn[i];
				string sMaleItem = info.malereward_itemid.ToString() + "," + info.malereward_itemcount.ToString() + "," + info.malereward_itemvalidity.ToString();
				string sFemaleItem = info.femalereward_itemid.ToString() + "," + info.femalereward_itemcount.ToString() + "," + info.femalereward_itemvalidity.ToString();
				if (info.HasSameTime(infoFirst))
				{
					strMale += CustomDefine.Separator_Item + sMaleItem;
					strFemale += CustomDefine.Separator_Item + sFemaleItem;
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

			strTemp = "function AddInTimeOnlineActivityData(index, value) " + Environment.NewLine 
			+ "\tif value ~= nil then" + Environment.NewLine 
			+ "\t\tlocal begintime = value[\"begintime\"];" + Environment.NewLine
			+ "\t\tlocal endtime = value[\"endtime\"];" + Environment.NewLine
			+ "\t\tlocal triggeringtime = value[\"triggeringtime\"];" + Environment.NewLine
			+ "\t\tlocal malereward = value[\"malereward\"];" + Environment.NewLine
			+ "\t\tlocal femalereward = value[\"femalereward\"];" + Environment.NewLine
			+ "\t\tlocal moneyreward = value[\"moneyreward\"];" + Environment.NewLine
			+ "\t\tlocal mailtitle = value[\"mailtitle\"];" + Environment.NewLine
			+ "\t\tlocal mailcontent = value[\"mailcontent\"];" + Environment.NewLine
			+ "\t\tAddInTimeOnlineActivity(index, begintime, endtime, triggeringtime, malereward, femalereward, moneyreward, mailtitle, mailcontent);" + Environment.NewLine
			+ "\tend" + Environment.NewLine 
			+ "end" + Environment.NewLine 
			+ Environment.NewLine 
		    + "function AddInTimeOnlineBriefInfo(weight)" + Environment.NewLine
			+ "\tif weight ~= nil then" + Environment.NewLine
			+ "\t\tAddInTimeOnlineActivityBriefInfo(exhibit, weight, regularImageName, thumbnailName, activity_title, activity_content, show_time_begin, show_time_end, activity_begin_announce, activity_end_announce);" + Environment.NewLine
			+ "\tend" + Environment.NewLine 
			+ "end" + Environment.NewLine 
			+ Environment.NewLine 
			+ "table.foreach(InTimeOnlineActivityData, AddInTimeOnlineActivityData);" + Environment.NewLine;
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
			public string show_time_begin { get; set; }
			[ExcelHeader("显示结束时间")]
			public string show_time_end { get; set; }
			[ExcelHeader("开始公告")]
			public string activity_begin_announce { get; set; }
			[ExcelHeader("结束公告")]
			public string activity_end_announce { get; set; }

			public void CorrectData()
			{
				if (null == regularImageName)
				{
					regularImageName = "";
				}
				if (null == thumbnailName)
				{
					thumbnailName = "";
				}
				if (null == activity_title)
				{
					activity_title = "";
				}
				if (null == activity_content)
				{
					activity_content = "";
				}
				if (null == show_time_begin)
				{
					show_time_begin = "";
				}
				if (null == show_time_end)
				{
					show_time_end = "";
				}
				if (null == activity_begin_announce)
				{
					activity_begin_announce = "";
				}
				if (null == activity_end_announce)
				{
					activity_end_announce = "";
				}
			}

		}

		class IntimeOnlineInfo
		{
			[ExcelHeader("日期")]
			public string activity_day { get; set; }
			[ExcelHeader("开始时间")]
			public string begintime { get; set; }
			[ExcelHeader("结束时间")]
			public string endtime { get; set; }
			[ExcelHeader("触发时间")]
			public string triggeringtime { get; set; }
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
			[ExcelHeader("邮件标题")]
			public string mailtitle { get; set; }
			[ExcelHeader("邮件内容")]
			public string mailcontent { get; set; }

			public void CorrectData()
			{
				if (null == activity_day)
				{
					activity_day = "";
				}
				if (null == begintime)
				{
					begintime = "";
				}
				if (null == endtime)
				{
					endtime = "";
				}
				if (null == triggeringtime)
				{
					triggeringtime = "";
				}
				if (null == malereward_itemname)
				{
					malereward_itemname = "";
				}
				if (null == femalereward_itemname)
				{
					femalereward_itemname = "";
				}
				if (null == mailtitle)
				{
					mailtitle = "";
				}
				if (null == mailcontent)
				{
					mailcontent = "";
				}
			}

			public bool HasSameTime(IntimeOnlineInfo info)
			{
				if (null == info)
				{
					return false;
				}
				if ( 0 != activity_day.CompareTo( info.activity_day ) )
				{
					return false;
				}
				if (0 != begintime.CompareTo(info.begintime))
				{
					return false;
				}
				if (0 != endtime.CompareTo(info.endtime))
				{
					return false;
				}
				if (0 != triggeringtime.CompareTo(info.triggeringtime))
				{
					return false;
				}

				return true;
			}
			
			public string BuildLua(int nCount, string strMale, string strFemale)
			{

				return "\t[" + nCount.ToString() + "] = { begintime = \"" + activity_day + " " + begintime
					+ "\", endtime = \"" + activity_day + " " + endtime
					+ "\", triggeringtime = \"" + triggeringtime
					+ "\", malereward = \"" + strMale + "\", femalereward = \"" + strFemale + "\", moneyreward = " + moneyreward
					+ "\", mailtitle = \"" + mailtitle
					+ "\", mailcontent = \"" + mailcontent
					+ "\" },";
			}

		}
	}
}
