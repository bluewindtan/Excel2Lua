--是否显示在主页, 0表示不显示，大于0表示显示 
local exhibit = 1;

--活动图片名称 
local regularImageName = "activity-ad_02";
local thumbnailName = "huodongzhongxin_02";

--活动名称 
local activity_title = "累计在线 天降好礼";
--活动内容，其中时间程序会自动填充 
local activity_content = "只要在%02d月%02d日 %02d:%02d~%02d月%02d日 %02d:%02d 活动期间，玩家游戏累计在线一定时长，即可获得奖励！";

--活动显示时间 
local show_begin_time = "2015-06-20 00:00:00";
local show_end_time = "2015-07-11 23:59:59";

--每日计时开始时间 
local resettime = {hour = 0, minute = 0, second = 0};

--玩家物品，activity_end_time与下一个activity_start_time不能完全相同
local onlineRewardInfo = 
{
	[1] = { activity_start_time = "2015-06-20 00:00:00", activity_end_time = "2015-06-20 23:59:59", detail = {
		[1] = {onlinetime = 900, malereward = "31026,1,-1", femalereward = "31026,1,-1", money = 288},
		[2] = {onlinetime = 1800, malereward = "31186,2,-1 | 31187,3,-1 | 31188,4,-1", femalereward = "31186,2,-1 | 31187,3,-1 | 31188,4,-1", money = 388},
		[3] = {onlinetime = 2700, malereward = "31189,1,86400", femalereward = "31189,1,86400", money = 688},
	}},
	[2] = { activity_start_time = "2015-06-27 00:00:00", activity_end_time = "2015-06-27 23:59:59", detail = {
		[1] = {onlinetime = 900, malereward = "31026,1,-1", femalereward = "31026,1,-1", money = 288},
		[2] = {onlinetime = 1800, malereward = "31186,1,-1", femalereward = "31186,1,-1", money = 388},
		[3] = {onlinetime = 3600, malereward = "31186,1,-1", femalereward = "31186,1,-1", money = 688},
	}},
}

function AddOnlineRewardActivityInfo(weight)
	local resetHour = resettime["hour"];
	local restMinute = resettime["minute"];
	local resetSeconds = resettime["second"];

	AddOnlineTimeRewardBrief(weight, exhibit, show_begin_time, show_end_time, regularImageName, thumbnailName, activity_title, activity_content, resetHour, restMinute, resetSeconds);
end

function AddOnlineReward(index, value, activityid, activity_start_time, activity_end_time) 
	if value ~= nil then
		local onlinetime = value["onlinetime"];	
		local malereward = value["malereward"];
		local femalereward = value["femalereward"];
		local money = value["money"];

		AddOnlineRewardInfo(activityid, activity_start_time, activity_end_time, index, onlinetime, malereward, femalereward, money);
	end
end

function AddOnlineRewardLoop(index, value)
	local detail = value["detail"];
	local activity_start_time = value["activity_start_time"];
	local activity_end_time = value["activity_end_time"];
	for i, v in ipairs(detail) do
		AddOnlineReward(i, v, index, activity_start_time, activity_end_time);
	end
end

table.foreach(onlineRewardInfo, AddOnlineRewardLoop);

