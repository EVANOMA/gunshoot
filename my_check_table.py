#-*-coding:utf-8 -*-
import period_recharge_reward_table
import period_recharge_sum_table
import importlib
# 首先运行Hash模块,如果包裹表有变动，需要重新运行
import file_content
from my_excel import *

# 参数区
# the last count of reward_table
a = 43

sum_data = period_recharge_sum_table.data[a]["Payment"]
reward_data = period_recharge_reward_table.data
ErrorMessage = []# 存放不一致的表项

row = 1
col = 1
fill_num = [0, 0, 0]

count = 1
for price, reward_value in reward_data.items():
    # TODO：price无法对应？
    text = ""
    Message = {}
    diaoluo = importlib.import_module(reward_value["RechargeAwardIndex"])

    row = headline(row, col, price, reward_value["RechargeAwardIndex"]) + 1
    if count == 1:
        adjustment()
        count = count + 1

    # 自我长度比对
    if len(reward_value["DisplayItem"]) != len(reward_value["DisplayItemCnt"]):
        fill_num[0] = 1
    if len(sum_data[price]["itemid"]) != len(sum_data[price]["itemcnt"]):
        fill_num[1] = 1
    if fill_num[0] == 1 or fill_num[1] == 1:
        pass
    else:
        if (len(reward_value["DisplayItem"]) == len(sum_data[price]["itemid"]) == len(diaoluo.data)):
            lenth = len(diaoluo.data)
        else:
            lenth = min(len(reward_value["DisplayItem"]), len(sum_data[price]["itemid"]), len(diaoluo.data))
            fill_num[0] = 2
            fill_num[1] = 2

        i = 0
        while (i < lenth):
            # TODO：检查元素类型是否一致？
            # 比对ID是否一致
            if reward_value["DisplayItem"][i] == sum_data[price]["itemid"][i] == diaoluo.data[i + 1]["ItemId"]:
                pass
            else:
                ErrorMessage.append(row + i)
                ErrorMessage.append(2)
            # 比对物品个数是否一致
            if reward_value["DisplayItemCnt"][i] == sum_data[price]["itemcnt"][i] == diaoluo.data[i + 1]["CountOneTime"]:
                pass
            else:
                ErrorMessage.append(row + i)
                ErrorMessage.append(3)
            i = i + 1

    # 写入reward_table
    write_s(0, reward_value["content"], row, 1)
    write(fill_num[0], reward_value["DisplayItem"], row, 2)
    write(fill_num[0], reward_value["DisplayItemCnt"], row, 3)
    # 写入sum_table
    # TODO:没有比对名字
    t = row
    for id in sum_data[price]["itemid"]:
        write_s(0, file_content.dic[id], t, 4)
        t = t + 1
    write(fill_num[1], sum_data[price]["itemid"], row, 5)
    write(fill_num[1], sum_data[price]["itemcnt"], row, 6)
    # 写入掉落表
    l_name = []
    l_id = []
    l_num = []
    for value in diaoluo.data.values():
        l_name.append(value["Name"])
        l_id.append(value["ItemId"])
        l_num.append(value["CountOneTime"])
        # TODO:掉落表自我比对
        write(0, l_name, row, 7)
        write(fill_num[2], l_id, row, 8)
        write(fill_num[2], l_num, row, 9)

    # TODO:没有考虑掉落表的长度
    length = max(len(reward_value["DisplayItem"]), len(reward_value["DisplayItemCnt"]),
                 len(sum_data[price]["itemid"]), len(sum_data[price]["itemcnt"]))
    row = row + length + 3

length = len(ErrorMessage)
for i in range(0, length, 2):
    mark(ErrorMessage[i], ErrorMessage[i+1])