--活动名称和内容 
local activityName = "新手七天乐";
local activityContent = "新玩家注册7天内，只要登陆就有好礼。累计登陆7天，可获得神秘大礼，千万别错过哦。";

--背包已满时，通过邮件发送奖励，邮件标题和内容 
local mailTitle = "奖励邮件标题";
local mailContent = "奖励邮件内容";

--活动时长, 单位：天 
local activityTime = 7;

--N天内的奖励配置，第几天及对应的奖励 
local activityReward = 
{
	[1] = { day = 1, maleItem = "32001,2,-1 | 14061,1,-1 | 33003,2,-1", femaleItem = "32001,2,-1 | 14593,1,-1 | 33003,2,-1", money = 388, bindCoin = 0,
			vipMaleItem = "32001,2,-1 | 337,1,1296000 | 33003,2,-1", vipFemaleItem = "32001,2,-1 | 937,1,1296000 | 33003,2,-1", vipMoney = 388, vipBindCoin = 0 };
	[2] = { day = 2, maleItem = "32001,2,-1", femaleItem = "32001,2,-1", money = 588, bindCoin = 0,
			vipMaleItem = "32001,2,-1", vipFemaleItem = "32001,2,-1", vipMoney = 588, vipBindCoin = 0 };
	[3] = { day = 3, maleItem = "32001,2,-1", femaleItem = "32001,2,-1", money = 888, bindCoin = 0,
			vipMaleItem = "32001,2,-1", vipFemaleItem = "32001,2,-1", vipMoney = 888, vipBindCoin = 0 };
	[4] = { day = 4, maleItem = "32001,2,-1", femaleItem = "32001,2,-1", money = 1088, bindCoin = 0,
			vipMaleItem = "32001,2,-1", vipFemaleItem = "32001,2,-1", vipMoney = 1088, vipBindCoin = 0 };
	[5] = { day = 5, maleItem = "32001,2,-1", femaleItem = "32001,2,-1", money = 1288, bindCoin = 0,
			vipMaleItem = "32001,2,-1", vipFemaleItem = "32001,2,-1", vipMoney = 1288, vipBindCoin = 0 };
	[6] = { day = 6, maleItem = "32001,2,-1", femaleItem = "32001,2,-1", money = 1588, bindCoin = 0,
			vipMaleItem = "32001,2,-1", vipFemaleItem = "32001,2,-1", vipMoney = 1588, vipBindCoin = 0 };
	[7] = { day = 7, maleItem = "32001,2,-1", femaleItem = "32001,2,-1", money = 1888, bindCoin = 0,
			vipMaleItem = "32001,2,-1", vipFemaleItem = "32001,2,-1", vipMoney = 1888, vipBindCoin = 0 };
};

--补签需要的金券 
local recvAgainMoney = 500;

--在线多少天可以获得全签奖励, 单位：天 
local cumulationTime = 7;

--全签奖励配置 
local cumulationReward = { maleItem = "32001,2,-1 | 8001,1,-1 | 31003,1,-1", femaleItem = "32001,2,-1 | 8001,1,-1 | 31003,1,-1", money = 2000, bindCoin = 0,
			vipMaleItem = "32001,2,-1 | 9050,1,1296000 | 31026,5,-1", vipFemaleItem = "32001,2,-1 | 9050,1,1296000 | 31026,5,-1", vipMoney = 2000, vipBindCoin = 0 };
AddFresherActivity(activityName, activityContent, mailTitle, mailContent, recvAgainMoney, activityTime, cumulationTime);
AddFresherActivityCumulationReward(cumulationReward.maleItem, cumulationReward.femaleItem, cumulationReward.money, cumulationReward.bindCoin,
		cumulationReward.vipMaleItem, cumulationReward.vipFemaleItem, cumulationReward.vipMoney, cumulationReward.vipBindCoin);

function AddFresherActivityRewardRoutine(index, value)
	if value ~= nil then
		AddFresherActivityReward(value.day, value.maleItem, value.femaleItem, value.money, value.bindCoin,
			value.vipMaleItem,value.vipFemaleItem,value.vipMoney,value.vipBindCoin);
	end
end

table.foreach(activityReward, AddFresherActivityRewardRoutine);


