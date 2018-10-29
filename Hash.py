import os
import importlib
import write


def file_name(path, Prefix):
    dic = {}
    start = len(path)
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        # 子目录
        if os.path.isdir(file_path):
            pass
        elif os.path.splitext(file_path)[0][start:start+len(Prefix)] == Prefix:
            s = os.path.splitext(file_path)[0][start:]
            py = importlib.import_module(s)
            try:
                for key, value in py.data.items():
                    dic[key] = value["Name"]
            except:
                pass

    write.write_py(dic)


file_dir = 'D:\\daobiao\\'
Prefix = 'package_item_table'
file_name(file_dir, Prefix)