--是否显示在主页, 0表示不显示，大于0表示显示 
local exhibit = 1;

-- 统计消费类型(1:M币 2:绑定M币和M币) 
local currencyType = 1;

--活动图片名称 
local regularImageName = "activity-ad_18";
local thumbnailName = "huodongzhongxin_18";

--活动名称 
local activity_title = "夏日物语，粉红恋人秀真爱！";
--活动内容 
local activity_content = "活动期间，在游戏中累计消费，即可成为粉红恋人，一起秀出火热的爱情！";

-- 活动起始时间 
local activity_begin_time = "2015-7-1 10:00:00";
local activity_end_time  = "2015-7-14 23:59:59";

-- 累计消费起始时间 
local spend_begin_time = "2015-7-1 10:00:00";
local spend_end_time = "2015-7-14 23:59:59";

--活动起始公告 
local spend_begin_annouce = "累计消费活动开始";
local spend_end_annouce = "累计消费活动结束";

local CummulativeSpendInfo = 
{
	[1] = { requirenum = 3000, malereward = "7051,1,-1", femalereward = "7051,1,-1", money = 0, bindmcoin = 0 },
	[2] = { requirenum = 10000, malereward = "31003,1,-1 | 33004,1,-1", femalereward = "31003,1,-1 | 33004,1,-1", money = 0, bindmcoin = 0 },
	[3] = { requirenum = 30000, malereward = "8110,1,604800 | 31014,1,-1", femalereward = "8110,1,604800 | 31014,1,-1", money = 0, bindmcoin = 0 },
	[4] = { requirenum = 50000, malereward = "8106,1,-1", femalereward = "8106,1,-1", money = 0, bindmcoin = 0 },
	[5] = { requirenum = 100000, malereward = "40238,1,-1 | 54238,1,-1", femalereward = "40738,1,-1 | 54738,1,-1", money = 0, bindmcoin = 0 },
	[6] = { requirenum = 200000, malereward = "42238,1,-1 | 44238,1,-1", femalereward = "57738,1,-1", money = 0, bindmcoin = 0 },
}

function AddCummulativeSpendTableInfo(index, value)
	if value ~= nil then
		local requirenum = value["requirenum"];
		local malereward = value["malereward"];
		local femalereward = value["femalereward"];
		local money = value["money"];
		local bindmcoin = value["bindmcoin"];

		AddCumulativeSpendRewardInfo(index, requirenum, malereward, femalereward, money, bindmcoin);
	end
end

function AddCumulativeSpend(weight)
	if weight ~= nil then
		AddCumulativeSpendInfo(exhibit, currencyType, weight, regularImageName, thumbnailName, activity_title, activity_content, activity_begin_time, activity_end_time, spend_begin_time, spend_end_time,spend_begin_annouce, spend_end_annouce);
		table.foreach(CummulativeSpendInfo, AddCummulativeSpendTableInfo);
	end
end


