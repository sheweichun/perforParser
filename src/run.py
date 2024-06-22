from calculate import calcalute_front



def show(data):
    print("====================")
    for row in data:
        name = row.get("name")
        total = row.get("total")
        label = row.get("cur_label")
        level = row.get("cur_level")
        print(f"row is {name} ----- {total}  {label} {level}")





if __name__ == "__main__":
    data = calcalute_front()

    show(data[0])
    show(data[1])
    show(data[2])
    show(data[3])



    # writeExcel(data)