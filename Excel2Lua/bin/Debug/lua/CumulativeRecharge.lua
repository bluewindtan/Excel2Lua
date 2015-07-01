--是否显示在主页, 0表示不显示，大于0表示显示 
local exhibit = 1;

--活动图片名称 
local regularImageName = "activity-ad_17";
local thumbnailName = "huodongzhongxin_17";

--活动名称 
local activity_title = "仲夏夜累充豪礼第二弹";
--活动内容 
local activity_content = "恋舞仲夏充值豪礼第二弹，绑定MB、海量英雄券大放送！";

-- 活动起始时间 
local activity_begin_time = "2015-7-1 10:00:00";
local activity_end_time  = "2015-7-14 23:59:59";

-- 累计充值起始时间 
local recharge_begin_time = "2015-7-1 10:00:00";
local recharge_end_time = "2015-7-14 23:59:59";

--活动起始公告 
local recharge_begin_annouce = "累计充值活动开始";
local recharge_end_annouce = "累计充值活动结束";

local CummulativeRechargeInfo = 
{
	[1] = { requirenum = 2000, malereward = "9077,1,-1", femalereward = "9077,1,-1", money = 0, bindmcoin = 2000 },
	[2] = { requirenum = 10000, malereward = "31005,1,-1", femalereward = "31005,1,-1", money = 0, bindmcoin = 3000 },
	[3] = { requirenum = 50000, malereward = "31187,5,-1", femalereward = "31187,5,-1", money = 0, bindmcoin = 10000 },
	[4] = { requirenum = 100000, malereward = "31187,10,-1", femalereward = "31187,10,-1", money = 0, bindmcoin = 20000 },
}

function AddCummulativeRechargeTableInfo(index, value)
	if value ~= nil then
		local requirenum = value["requirenum"];
		local malereward = value["malereward"];
		local femalereward = value["femalereward"];
		local money = value["money"];
		local bindmcoin = value["bindmcoin"];

		AddCummulativeRecharge(index, requirenum, malereward, femalereward, money, bindmcoin );
	end
end

function AddCumulativeInfo(weight)
	if weight ~= nil then
		AddCumulativeBrief(exhibit, weight, regularImageName, thumbnailName, activity_title, activity_content, activity_begin_time, activity_end_time, recharge_begin_time, recharge_end_time, recharge_begin_annouce, recharge_end_annouce);
		table.foreach(CummulativeRechargeInfo, AddCummulativeRechargeTableInfo);
	end
end


