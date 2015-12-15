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
		private Dictionary< uint, ItemInfo > m_dicItem = null;

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
			m_dicItem = new Dictionary< uint, ItemInfo >();
			ExcelReader reader = new ExcelReader(sFile);
			List<ItemInfo> itemList = reader.GetSheetData<ItemInfo>("道具");
			List<ItemInfo> clothList = reader.GetSheetData<ItemInfo>("服饰");
			List<ItemInfo> oldEffectClothList = reader.GetSheetData<ItemInfo>("旧特效服饰");
			List<ItemInfo> badgeList = reader.GetSheetData<ItemInfo>("徽章");
			List<ItemInfo> externItemList = reader.GetSheetData<ItemInfo>("扩展物品");
			reader.Close();

			foreach (ItemInfo item in itemList)
			{
				m_dicItem.Add(item.Type, item);
			}
			foreach (ItemInfo item in clothList)
			{
				item.bOldEffectCloth = false;
				m_dicItem.Add(item.Type, item);
            }
			foreach (ItemInfo item in oldEffectClothList)
			{
				item.bOldEffectCloth = true;
				m_dicItem.Add(item.Type, item);
			}			
			foreach (ItemInfo item in badgeList)
			{
				m_dicItem.Add(item.Type, item);
			}
			foreach (ItemInfo item in externItemList)
			{
				m_dicItem.Add(item.Type, item);
			}
		}

		public ItemInfo GetItem(uint nItemID)
		{
			if (m_dicItem.ContainsKey(nItemID))
			{
				return m_dicItem[nItemID];
			}

			return null;
		}

		public bool IsHaveItem(uint nItemID)
		{
			return m_dicItem.ContainsKey(nItemID);
		}

		public bool CheckItemAndLogError(string strWhereAction, uint nItemID, int nItemCount, int nItemValidity)
		{
			// check if item exists
			if (!IsHaveItem(nItemID))
			{
				// log 
				string strMsg = "Not exists Item=" + nItemID.ToString();
				LogMgr.Instance.WriteLog(strWhereAction + " : " + strMsg);
				return false;
			}

			// get item 
			ItemInfo item = GetItem(nItemID);
			if (item != null)
			{
				// check if cloth or badge 
				if (item.IsCloth() || item.IsBadge())
				{
					if (nItemCount != 1)
					{
						// log 
						string strMsg = "Item=" + nItemID.ToString() + " is cloth or badge, so the count must be 1, not be " + nItemCount.ToString();
						LogMgr.Instance.WriteLog(strWhereAction + " : " + strMsg);
						return false;
					}
				}
			}

			return true;
		}
	
	}
	
}
