from myparser import parseExcel, extract_data, writeExcel
from os import path
def rule_quota(data, quota_list):
    quota_index = 0
    new_rank_data = []
    for row in data:
        if quota_index >= len(new_rank_data):
            cur_rank_list = []
            new_rank_data.append(cur_rank_list)
        cur_rank_list = new_rank_data[quota_index]
        if quota_index >= len(quota_list):
            cur_rank_list.append(row)
            continue
        cur_quota = quota_list[quota_index]
        tmp_num = cur_quota.get("tmp_num")
        row["cur_label"] = cur_quota.get("label")
        row["cur_level"] = cur_quota.get("level")
        cur_rank_list.append(row)
        new_tmp_num = tmp_num - 1
        cur_quota["tmp_num"] = new_tmp_num
        if new_tmp_num == 0:
            quota_index = quota_index + 1
        # if quota_index >= len(quota_list):
        #     break
    return new_rank_data
        
def rule_two_quota(data, quota_list):
    cur_for_level = 1
    cur_for_index = 0
    for rank_list in data:
        rank_for_index = 0
        for item in rank_list:
            last_level = item.get("last_level")
            cur_level = item.get("cur_level")
            name = item.get("name")
            if cur_level - last_level > 2: # 降得太快了
                # print(f"降得太快了, {name}, {cur_level}, ${cur_for_level}, {last_level}")
                # print(f"c -> l  > 2, {cur_level}, ${cur_for_level}, {last_level}, {name}")
                last_rank_list = data[cur_level - 2]
                last_rank_index = len(last_rank_list) - 1
                last_rank_item = last_rank_list[last_rank_index]
                last_rank_list[last_rank_index] = item
                my_sort(last_rank_list)
                new_cur_level = cur_level - 1
                item["cur_level"] = new_cur_level
                item["cur_label"] = get_quato_label(new_cur_level, quota_list)

                last_rank_item["cur_level"] = cur_level
                last_rank_item["cur_label"] = get_quato_label(cur_level, quota_list)

                del rank_list[rank_for_index]
                rank_list.insert(0, last_rank_item)
                my_sort(rank_list)
                return rule_two_quota(data, quota_list)
            elif last_level - cur_level > 2: #升得太快了
                # print(f"升得太快了, {name}, {cur_level}, ${cur_for_level}, {last_level}")
                next_rank_list = data[cur_level]
                next_rank_item = next_rank_list[0]
                next_rank_list[0] = item
                my_sort(next_rank_list)
                new_cur_level = cur_level + 1
                item["cur_level"] = new_cur_level
                item["cur_label"] = get_quato_label(new_cur_level, quota_list)

                next_rank_item["cur_level"] = cur_level
                next_rank_item["cur_label"] = get_quato_label(cur_level, quota_list)

                del rank_list[rank_for_index]
                rank_list.append(next_rank_item)
                my_sort(rank_list)
                return rule_two_quota(data, quota_list)
                # print(f"l -> c  > 2, {cur_level}, {last_level}, {name}")
            rank_for_index = rank_for_index + 1
        cur_for_level = cur_for_level + 1
        cur_for_index = cur_for_index + 1


        
        
    

def my_sort(data):
    data.sort(key=lambda person: person["total"], reverse=True)
    return data

def is_None(val):
    if val is None:
        return True
    elif isinstance(val, str):
            newVal = val.replace(" ", "")
            return newVal == ""
    return False
    
def get_quato_level(label, quato_list):
    for quato in quato_list:
        quato_label = quato.get("label")
        if quato_label == label:
            return quato.get("level")

def get_quato_label(level, quato_list):
    index = level - 1
    if index < len(quato_list):
        item = quato_list[index]
        return item.get("label")

def calcalute_front():
    quota_list = [
        {"label": '公司核心', "number": 20, "level": 1},
        {"label": '公司优秀', "number": 30, "level": 2},
        {"label": '部门优秀', "number": 50, "level": 3}
    ]
    return calcalute_result("frontend1", quota_list)

def one2Two(data, quota_list):
    valid_data = []
    invalid_data = []
    max_level = len(quota_list) + 1
    for item in data: #把无效的数据都提取出来
        leave_time = item.get('leave_time')
        is_cancel = item.get('is_cancel')
        is_promote = item.get('is_promote')
        # name = item.get("name")
        # if name == "友516":
            # print(f"{name}, {leave_time}, {is_cancel}, {is_promote}")
        if is_None(leave_time) and is_None(is_cancel) and is_None(is_promote):
            last_level_label = item.get("last_level")
            last_level = get_quato_level(last_level_label, quota_list)
            # print(f"last_level : oldval is {last_level_label}, {last_level} is { last_level is None}")
            item["cur_level"] = max_level
            if last_level is None:
                item["last_level"] = max_level
            else:
                item["last_level"] = last_level
            valid_data.append(item)
        else:
            invalid_data.append(item)
    return {
        "valid_data": valid_data,
        "invalid_data": invalid_data
    }

def calcalute_result(name, quota_list):
    for quota in quota_list:
        quota["tmp_num"] = quota.get("number")
    current_file_path = path.realpath(__file__)
    current_dir = path.dirname(current_file_path)
    docsDir = path.join(current_dir, "..", "docs")
    data = parseExcel(path.join(docsDir, f"{name}.xlsx"), extract_data)
    result = one2Two(data, quota_list)
    valid_data = result.get('valid_data')
    # invalid_data = result.get('invalid_data')
    my_sort(valid_data)
    rank_data = rule_quota(valid_data, quota_list)
    rule_two_quota(rank_data, quota_list)
    return rank_data