using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using MSExcelReader;

namespace Excel2Lua
{
	public class E2L_Packet
	{
		string m_strFullExcelName = "";
		string m_strDirector = "";
		string m_strExcel = "";

		ExcelReader m_excelReader = null;
		StreamWriter m_sWriter = null;

		List<cPacketInfo> m_listPacketInfo = new List<cPacketInfo>();

		public E2L_Packet(string strDirector, string strExcel)
		{
			m_strFullExcelName = strDirector + "\\" + strExcel;
			m_strDirector = strDirector;		
			m_strExcel = strExcel;
		}

		public bool ReadExcel()
		{
			if (null == m_excelReader)
			{
				m_excelReader = new ExcelReader(m_strFullExcelName);
			}
			// 
			int nSheets = m_excelReader.GetSheetsCount();
			for (int i = 0; i < nSheets; i++)
			{
				string strPacketID = m_excelReader.GetSheetNameByIndex(i);
				List<PacketInfo> listInfo = m_excelReader.GetSheetData<PacketInfo>(strPacketID);
				//
				cPacketInfo info = new cPacketInfo(strPacketID);
				info.CopyAllPacket(listInfo);
				m_listPacketInfo.Add(info);
			}

			return true;
		}

		public bool CheckData()
		{
			bool bReturn = true;
			foreach (cPacketInfo info in m_listPacketInfo)
			{
				if (!info.CheckData())
				{
					bReturn = false;
				}
			}

			return bReturn;
		}

		public bool SaveToLua()
		{
			// Create the lua director 
			string targetDir = m_strDirector + "\\..\\lua\\item";
			if (!Directory.Exists(targetDir))
			{
				Directory.CreateDirectory(targetDir);
			}

			// UTF8 NO BOM 
			foreach (cPacketInfo info in m_listPacketInfo)
			{
				// lua file 
				string strLua = targetDir + "\\" + CustomDefine.PACKET_LUA_NAME_PREFIX + info.PacketID + ".lua";
				m_sWriter = new StreamWriter(strLua, false, CustomDefine.WRITE_ENCODING);
				if (null != m_sWriter)
				{
					info.SaveData(m_sWriter);
					m_sWriter.Flush();
					m_sWriter.Close();
				}			
			}

			return true;
		}

		class PacketInfo
		{
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
			public int moneyreward { get; set; }
			
		}

		class cPacketInfo
		{
			string m_strID = "";
			List<PacketInfo> m_listReward = null;

			public cPacketInfo(string strID)
			{
				m_strID = strID;
			}

			public string PacketID
			{
				get
				{
					return m_strID;
				}
			}
			
			public void AddOnePacket(PacketInfo info)
			{
				if (m_listReward != null)
				{
					m_listReward.Add(info);
				}
			}

			public void CopyAllPacket(List<PacketInfo> listInfo)
			{
				m_listReward = new List<PacketInfo>(listInfo);
			}

			public void SaveData(StreamWriter sw)
			{
				// write to lua 
				sw.WriteLine();
				sw.WriteLine("-------------------以下是一个礼包的定义\r\n--礼品包的详细信息 ");
				string strTemp = "function GetPacketItemsInfo_" + PacketID + "(RoleIndex, PacketID)";
				sw.WriteLine(strTemp);
				sw.WriteLine("\tlocal Sex = GetRoleSex(RoleIndex)");
				sw.WriteLine("\tif Sex == SexType_Male then");
				//////////////////////////////////////////////////
				string strMaleReward = "\t\treturn \"";
				string strFemaleReward = "\t\treturn \"";
				string strMaleTable = string.Empty;
				string strFemaleTable = string.Empty;
				List<string> listMaleItem = new List<string>();
				List<string> listFemaleItem = new List<string>();
				string strMaleItem = string.Empty;
				string strFemaleItem = string.Empty;
				string strMaleAddItem = string.Empty;
				string strFemaleAddItem = string.Empty;
				int nMoneyReward = 0;
				int nItemCount = 0;
				foreach (PacketInfo info in m_listReward)
				{
					if (nMoneyReward <= 0)
					{
						nMoneyReward = info.moneyreward;
					}
					if (0 < nItemCount)
					{
						strMaleReward += CustomDefine.Separator_In_Item; // 这个脚本比较特殊：用的是，不是| 
						strFemaleReward += CustomDefine.Separator_In_Item;
						strMaleTable += Environment.NewLine;
						strFemaleTable += Environment.NewLine;
					}
					strMaleItem = info.malereward_itemid.ToString() 
						+ CustomDefine.Separator_In_Item + info.malereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.malereward_itemvalidity).ToString();
					strFemaleItem = info.femalereward_itemid.ToString() 
						+ CustomDefine.Separator_In_Item + info.femalereward_itemcount.ToString()
						+ CustomDefine.Separator_In_Item + CustomFunc.Day2Second(info.femalereward_itemvalidity).ToString();
					strMaleAddItem = "\t\tAddItemToTable(ItemTable,  " + strMaleItem + ")";
					strFemaleAddItem = "\t\tAddItemToTable(ItemTable,  " + strFemaleItem + ")";
					listMaleItem.Add(strMaleItem);
					listFemaleItem.Add(strFemaleItem);
					nItemCount++;
					strMaleReward += strMaleItem;
					strFemaleReward += strFemaleItem;
					strMaleTable += strMaleAddItem;
					strFemaleTable += strFemaleAddItem;
				}
				strMaleReward += "\"		--物品1ID 数量 有效时间1, 物品2ID 数量2 有效时间2, 物品ID3 数量3 有效时间3";
				strFemaleReward += "\"		--物品1ID 数量 有效时间1, 物品2ID 数量2 有效时间2, 物品ID3 数量3 有效时间3";
				sw.WriteLine(strMaleReward);
				sw.WriteLine("\telse");
				sw.WriteLine(strFemaleReward);
				//////////////////////////////////////////////////
				sw.WriteLine("\tend");
				sw.WriteLine("end");
				sw.WriteLine();

				sw.WriteLine("-- 礼包定义,需要填物品ID号,数量");
				strTemp = "function PacketGetReward_" + PacketID + "(RoleIndex)";
				sw.WriteLine(strTemp);
				strTemp = "\tlocal nPacketID = " + PacketID;
				sw.WriteLine(strTemp);
				sw.WriteLine("\tlocal Sex = GetRoleSex(RoleIndex)");
				sw.WriteLine("\t--礼包的道具数量,必填");
				strTemp = "\tlocal Count = " + nItemCount.ToString();
				sw.WriteLine(strTemp);
				sw.WriteLine("\tlocal ItemTable = {}");
				sw.WriteLine("\t--------------------------------ItemID    数量  有效期  ");
				sw.WriteLine("\tif Sex == SexType_Male then ");
				sw.WriteLine(strMaleTable);
				sw.WriteLine("\telse");
				sw.WriteLine(strFemaleTable);
				sw.WriteLine("\tend");
				sw.WriteLine();

				sw.WriteLine("\tif CanAddPacketItem(RoleIndex,nPacketID) == 1 then");
				sw.WriteLine("\t\tfor	i = 1, Count do");
				sw.WriteLine("\t\t\tAddItemToRole(RoleIndex,ItemTable[i][1],ItemTable[i][2], ItemTable[i][3],nPacketID);");
				sw.WriteLine("\t\tend");
				sw.WriteLine("\t\tModifyRoleMoney(RoleIndex, " + nMoneyReward.ToString() + ");	------第二个参数表示增加的金券数,如果需要添加金券,则需要修改该数字；没有该函数或该数字为0（这条语句）表示金券为0,");
				sw.WriteLine("\t\treturn 1");
				sw.WriteLine("\telse");
				sw.WriteLine("\t\treturn 0");
				sw.WriteLine("\tend");
				sw.WriteLine("end");
				sw.WriteLine();

				sw.WriteLine("--注册获取礼包信息函数");
				sw.WriteLine("-- 礼包ID 礼包描述函数");
				strTemp = "RegisterGetPacketInfo(" + PacketID + ", GetPacketItemsInfo_" + PacketID + ")";
				sw.WriteLine(strTemp);
				sw.WriteLine("--注册礼包使用函数");
				sw.WriteLine("--                    礼包ID 礼包获取函数");
				strTemp = "RegisterGetPacketItem(" + PacketID + ", PacketGetReward_" + PacketID + ")";
				sw.WriteLine(strTemp);
				sw.WriteLine();
			}
			
			public bool CheckData()
			{
				bool bReturn = true;
				foreach (PacketInfo info in m_listReward)
				{
					bool bCheckMale = ItemMgr.Instance.CheckItemAndLogError(CustomDefine.PACKET_EXCEL_NAME, info.malereward_itemid, info.malereward_itemcount, info.malereward_itemvalidity);
					bool bCheckFemale = ItemMgr.Instance.CheckItemAndLogError(CustomDefine.PACKET_EXCEL_NAME, info.femalereward_itemid, info.femalereward_itemcount, info.femalereward_itemvalidity);
					if (!bCheckFemale || !bCheckMale)
					{
						bReturn = false;
					}
				}
				return bReturn;
			}
		}
	}
}
