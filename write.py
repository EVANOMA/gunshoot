def write_py(dic):
    file = open('file_content.py', 'w')
    # file.write("#-*-coding:utf-8 -*-\n")
    file.write("dic=")
    file.write(str(dic))
    file.close()


