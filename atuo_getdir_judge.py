#python自动化办公学习记录
#shawan
#data：20200516
#对照网站内容学习自动化办公（1）对相对文件夹内文件进行检查判断，便于后续操作
import os
num = 0
files = os.listdir()


print('当前文件件下非文件分别为：')
for file in os.listdir():
    if os.path.isdir(file) == False:    #使用for遍历及os.path.isdir(file)逐个判断文件是否为文件夹
        print(file)
print("\n")

#获取文件夹中包含“python”单词的文件夹或文件，单词不区分大小写，逐个输出后统计总数
print('包含"python"单词的文件or文件夹分别是:')
for i in files:
    l_name = i.lower()
    if "python" in l_name:
        print(i)
        num = num + 1
    else:
        continue
print(f'总计个数为：{num}个')
