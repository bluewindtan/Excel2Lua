--是否显示在主页, 0表示不显示，大于0表示显示 
local exhibit = 1;

--活动图片名称 
local regularImageName = "activity-ad_11";
local thumbnailName = "huodongzhongxin_11";

--活动名称 
local activity_title = "准点在线礼品多";
--活动内容 
local activity_content = "只要在2015年6月21日~2015年7月12日活动期间，在以下时间点准时在线，即可获得丰厚的奖励！";

--活动显示起始时间 
local show_time_begin = "2015-06-21 00:00:00";
local show_time_end = "2015-07-12 23:59:59";

--活动起始公告 
local activity_begin_announce = "";
local activity_end_announce = "";

local InTimeOnlineActivityData = 
{
	[1] = { begintime = "2015-06-21 12:30:00", endtime = "2015-06-21 22:30:00", triggeringtime = "20:00", malereward = "31186,1,-1 | 31189,3,-1 | 31190,4,-1", femalereward = "31186,1,-1 | 31189,3,-1 | 31190,4,-1", moneyreward = 288, mailtitle = "20点在线奖励", mailcontent = "恭喜您，获得20点准点在线奖励，祝您游戏愉快！" },
	[2] = { begintime = "2015-06-21 12:30:00", endtime = "2015-06-21 22:30:00", triggeringtime = "21:00", malereward = "31187,2,-1", femalereward = "31187,2,-1", moneyreward = 488, mailtitle = "21点在线奖励", mailcontent = "恭喜您，获得21点准点在线奖励，祝您游戏愉快！" },
	[3] = { begintime = "2015-06-28 12:30:00", endtime = "2015-06-28 22:30:00", triggeringtime = "22:00", malereward = "31188,1,86400", femalereward = "31188,1,86400", moneyreward = 688, mailtitle = "22点在线奖励", mailcontent = "恭喜您，获得22点准点在线奖励，祝您游戏愉快！" },
}

function AddInTimeOnlineActivityData(index, value) 
	if value ~= nil then
		local begintime = value["begintime"];
		local endtime = value["endtime"];
		local triggeringtime = value["triggeringtime"];
		local malereward = value["malereward"];
		local femalereward = value["femalereward"];
		local moneyreward = value["moneyreward"];
		local mailtitle = value["mailtitle"];
		local mailcontent = value["mailcontent"];
		AddInTimeOnlineActivity(index, begintime, endtime, triggeringtime, malereward, femalereward, moneyreward, mailtitle, mailcontent);
	end
end

function AddInTimeOnlineBriefInfo(weight)
	if weight ~= nil then
		AddInTimeOnlineActivityBriefInfo(exhibit, weight, regularImageName, thumbnailName, activity_title, activity_content, show_time_begin, show_time_end, activity_begin_announce, activity_end_announce);
	end
end

table.foreach(InTimeOnlineActivityData, AddInTimeOnlineActivityData);

