using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcelReader;
using System.Text.RegularExpressions;

namespace Excel2Lua
{
    public class ItemInfo
    {
        [ExcelHeader("物品ID")]
        public ushort Type { get; set; }
        [ExcelHeader("物品名称")]
        public string Name { get; set; }
        [ExcelHeader("类别")]
        public byte AnType_1 { get; set; }
        [ExcelHeader("二级子类")]
        public byte AnType_2 { get; set; }
        [ExcelHeader("三级子类")]
        public byte AnType_3 { get; set; }
        [ExcelHeader("男/女")]
        public short SexNeed { get; set; }
        [ExcelHeader("有效时间")]
        public int Matune { get; set; }
        [ExcelHeader("使用次数")]
        public short UseTimes { get; set; }
        [ExcelHeader("等级要求")]
        public ushort NeedLevel { get; set; }
        [ExcelHeader("唯一性")]
        public bool Unique { get; set; }
        [ExcelHeader("叠加")]
        public uint MaxStackNumber { get; set; }
        [ExcelHeader("CD类型")]
        public ushort CoolDownTag { get; set; }
        [ExcelHeader("CD时间")]
        public int CoolDownTime { get; set; }
        [ExcelHeader("ICON")]
        public String Icon { get; set; }
        [ExcelHeader("Atlas")]
        public String Atlas { get; set; }
        [ExcelHeader("Animation")]
        public String Animation { get; set; }
        [ExcelHeader("VIP")]
        public ushort Vip { get; set; }
        [ExcelHeader("Intimacy")]
        public int Intimacy { get; set; }
        [ExcelHeader("TIPS")]
        public String Intro { get; set; }
        [ExcelHeader("新手道具")]
        public bool IsFresher { get; set; }

        [ExcelHeader("基因0")]
        public ushort GeneID0 { get; set; }
        [ExcelHeader("参数01")]
        public int ParamA0 { get; set; }
        [ExcelHeader("参数02")]
        public int ParamB0 { get; set; }
        [ExcelHeader("参数03")]
        public String Param0 { get; set; }
        [ExcelHeader("基因1")]
        public ushort GeneID1 { get; set; }
        [ExcelHeader("参数11")]
        public int ParamA1 { get; set; }
        [ExcelHeader("参数12")]
        public int ParamB1 { get; set; }
        [ExcelHeader("参数13")]
        public String Param1 { get; set; }
        [ExcelHeader("基因2")]
        public ushort GeneID2 { get; set; }
        [ExcelHeader("参数21")]
        public int ParamA2 { get; set; }
        [ExcelHeader("参数22")]
        public int ParamB2 { get; set; }
        [ExcelHeader("参数23")]
        public String Param2 { get; set; }
        [ExcelHeader("基因3")]
        public ushort GeneID3 { get; set; }
        [ExcelHeader("参数31")]
        public int ParamA3 { get; set; }
        [ExcelHeader("参数32")]
        public int ParamB3 { get; set; }
        [ExcelHeader("参数33")]
        public String Param3 { get; set; }
        [ExcelHeader("基因4")]
        public ushort GeneID4 { get; set; }
        [ExcelHeader("参数41")]
        public int ParamA4 { get; set; }
        [ExcelHeader("参数42")]
        public int ParamB4 { get; set; }
        [ExcelHeader("参数43")]
        public String Param4 { get; set; }
        [ExcelHeader("基因5")]
        public ushort GeneID5 { get; set; }
        [ExcelHeader("参数51")]
        public int ParamA5 { get; set; }
        [ExcelHeader("参数52")]
        public int ParamB5 { get; set; }
        [ExcelHeader("参数53")]
        public String Param5 { get; set; }

		// 道具
		public bool IsEquip()
		{
			return AnType_1 == (byte)EItemClassType.eItemClassType_Equip;
		}

		// 服饰
		public bool IsCloth()
		{
			return IsEquip() && (AnType_2 == (byte)EItemEquipType.eItemEquipType_Cloth);
		}

		// 徽章
		public bool IsBadge()
		{
			return IsEquip() && (AnType_2 == (byte)EItemEquipType.eItemEquipType_Badge);
		}

		/////////////////////////////////////////////////////////////////////
		//一级类型
		/////////////////////////////////////////////////////////////////////

		//装备，道具
		public enum EItemClassType
		{
			eItemClassType_None = 0,
			eItemClassType_Equip,       //装备
			eItemClassType_Expendable,  //道具(消耗品)
			eItemClassType_Max
		}


		////////////////////////////////////////////////////////////////////
		//二级类型
		////////////////////////////////////////////////////////////////////

		// 装备子类型
		public enum EItemEquipType
		{
			eItemEquipType_None = 0,
			eItemEquipType_Cloth,       //衣服和饰品
			eItemEquipType_Badge,       //徽章
			eItemEquipType_Max
		}

		// 消耗品子类型
		public enum EItemExpendableType
		{
			eItemExpendableType_None = 0,
			eItemExpendableType_Function,   //功能类
			eItemExpendableType_Social,     //社交类
			eItemExpendableType_Addition,   //增益类
			eItemExpendableType_Packet,     //礼包类
			eItemExpendableType_Box,        //宝箱类
			eItemExpendableType_TransForm,  //变身类
			eItemExpendableType_DynamicBox, //动态宝箱
			eItemExpendableType_PetEgg,     //宠物蛋
			eItemExpendableType_PetItem,    //宠物物品
			eItemExpendableType_Max
		}


		////////////////////////////////////////////////////////////////////
		// 三级类型
		////////////////////////////////////////////////////////////////////

		// 衣服子类型
		public enum EItemClothType
		{
			eItemClothType_Invalid = -1,
			eItemClothType_Hair,          //发型
			eItemClothType_Face,          //表情
			eItemClothType_Body,          //上衣
			eItemClothType_Glove,         //手套
			eItemClothType_Leg,           //下装
			eItemClothType_Cap,           //头饰
			eItemClothType_FacialContent, //脸饰
			eItemClothType_Shoulder,      //肩膀
			eItemClothType_Wing,          //翅膀
			eItemClothType_LeftHand,      //左手持
			eItemClothType_RightHand,     //右手持
			eItemClothType_Wrist,         //手腕
			eItemClothType_Hip,           //臀部
			eItemClothType_Sock,          //袜子
			eItemClothType_Feet,          //鞋子
			eItemClothType_Skin,          //肤色
			eItemClothType_All,           //全身装，除ItemCloth_Type_Skin
			eItemClothType_Suit,          //套装
			eItemClothType_Max,

		}

		// 徽章类型
		public enum EItemBadgeType
		{
			eItemBadgeType_WingdingRing  =  1,  //结婚戒指
			eItemBadgeType_SpecialRing,         //个性戒指
			eItemBadgeType_VIP,                 //VIP徽章
			eItemBadgeType_Show,                //展示特效
			eItemBadgeType_ExpEffect,           //经验效果
			eItemBadgeType_CardDecoration,      //名片装饰
		}
    }
}


