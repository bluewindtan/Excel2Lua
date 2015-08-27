using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;

namespace Excel2Lua
{
	public sealed class ItemMgr
	{
		private static ItemMgr instance = null;
		private static readonly object padlock = new object();
		private Dictionary< ushort, ItemInfo > m_dicItem = null;

		public static ItemMgr Instance
		{
			get
			{
				if (null == instance)
				{
					lock (padlock)
					{
						if (null == instance)
						{
							instance = new ItemMgr();
						}
					}
				}
				return instance;
			}
		}

		public void ReadExcelItem(string sFile)
		{
			if (null != m_dicItem)
			{
				return;
			}
			m_dicItem = new Dictionary< ushort, ItemInfo >();
			ExcelReader reader = new ExcelReader(sFile);
			List<ItemInfo> itemList = reader.GetSheetData<ItemInfo>("道具");
			List<ItemInfo> clothList = reader.GetSheetData<ItemInfo>("服饰");
			List<ItemInfo> badgeList = reader.GetSheetData<ItemInfo>("徽章");
			List<ItemInfo> externItemList = reader.GetSheetData<ItemInfo>("扩展物品");
			reader.Close();

			foreach (ItemInfo info in itemList)
			{
				m_dicItem.Add(info.Type, info);
			}
			foreach (ItemInfo info in clothList)
			{
				m_dicItem.Add(info.Type, info);
			}
			foreach (ItemInfo info in badgeList)
			{
				m_dicItem.Add(info.Type, info);
			}
			foreach (ItemInfo info in externItemList)
			{
				m_dicItem.Add(info.Type, info);
			}
		}

		public ItemInfo GetItem(ushort nItemID)
		{
			if (m_dicItem.ContainsKey(nItemID))
			{
				return m_dicItem[nItemID];
			}

			return null;
		}

		public bool IsHaveItem(ushort nItemID)
		{
			return m_dicItem.ContainsKey(nItemID);
		}

		public bool CheckItemAndLogError(ushort nItemID, string strWhereAction)
		{
			if (IsHaveItem(nItemID))
			{
				return true;
			}
			// log 
			string strMsg = "Not exists Item=" + nItemID.ToString();
			LogMgr.Instance.WriteLog(strWhereAction + " : " + strMsg);
			return false;
		}
	
	}
	
}
