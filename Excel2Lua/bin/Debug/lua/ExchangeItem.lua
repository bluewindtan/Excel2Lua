--是否显示在主页, 0表示不显示，大于0表示显示 
local exhibit = 1;

--活动图片名称 
local regularImageName = "activity-ad_09";
local thumbnailName = "huodongzhongxin_09";

--活动显示起始时间 
local beginExhibitTime = "2015-6-17 00:00:00";
local endExhibitTime = "2015-6-30 23:59:59";

--活动起始时间 
local exchangeBeginTime = "2015-6-17 00:00:00";
local exchangeEndTime = "2015-6-30 23:59:59";

--活动名称 
local activity_Title = "英雄大秀场之呆绿酷雷";

--活动内容 
local exchangeActivityIntro = "活动期间收集萌动英雄券A，即有机会获得恋舞超萌英雄和各类精美装饰哦~";

--活动起始公告 
local exchangeStartAnnouce = "";
local exchangeEndAnnouce = "";

--兑换物品 
local exchangeItemType = 31186;

local ExhangeRewardTable = 
{
	[1] = { requireItemNum = 6, maleReward = "31184,1,-1", femaleReward = "31184,1,-1", money = 288 },
	[2] = { requireItemNum = 10, maleReward = "35195,1,-1", femaleReward = "35195,1,-1", money = 388 },
	[3] = { requireItemNum = 60, maleReward = "40211,1,-1 | 54211,1,-1", femaleReward = "40702,1,-1 | 54702,1,-1", money = 488 },
	[4] = { requireItemNum = 100, maleReward = "42211,1,-1 | 44211,1,-1", femaleReward = "57702,1,-1", money = 588 },
	[5] = { requireItemNum = 200, maleReward = "8112,1,-1", femaleReward = "8112,1,-1", money = 688 },
}

function AddExchangeItemTableInfo(index, value)
	if value ~= nil then
		local requireItemNum = value["requireItemNum"];
		local maleReward = value["maleReward"];
		local femaleReward = value["femaleReward"];
		local money = value["money"];
		AddExchangeItemReward(index, requireItemNum, maleReward, femaleReward, money);
	end
end

function AddExchangeItemInfo(weight)
	if weight ~= nil then
		AddExchangeItemBriefInfo(exhibit, weight,regularImageName, thumbnailName, exchangeItemType, beginExhibitTime, endExhibitTime, exchangeBeginTime, exchangeEndTime, activity_Title, exchangeActivityIntro, exchangeStartAnnouce, exchangeEndAnnouce);
		table.foreach(ExhangeRewardTable, AddExchangeItemTableInfo);
	end
end

