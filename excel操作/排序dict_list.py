def sortedDictValues1(adict):
    items = adict.items()
    print(items)
    items.sort()
    return [value for key, value in items]


def sortedDictValues2(adict):
    keys = adict.keys()
    print(type(keys))
    # print('tttttt', keys)
    # keys.sort()
    return [adict[key] for key in keys]


def sort_value(adict):
    for k in sorted(adict, key=adict.__getitem__):
        print(k, adict[k])
    return [adict[key] for key in sorted(adict, key=adict.__getitem__)]


def sort_key(adict):
    for k in sorted(adict):
        print(k, adict[k])
    pass


data = {'2': 3, '1': 5, '6': 1, '5': 1, '0': 1}
print("value:", sort_value(data))
sort_value(data)
print(sorted(data))
# print(out_data)
