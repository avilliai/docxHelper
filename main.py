# 导入模块
import os

import yaml

import docx

rep = {}
def checkD():
    while True:
        print("-------执行关键词管理-------")
        order=input("按1开始替换\n按2删除替换词\n按任意键添加替换词\n输入:")
        if order=="1":
            print("当前替换任务词典:"+str(rep))
            break
        if order=="2":
            delW=input("输入要删除的替换词(原词):")
            rep.pop(delW)
            print("当前替换任务词典:" + str(rep))
            with open("replace.yaml", 'w', encoding="utf-8") as file:
                yaml.dump(rep, file, allow_unicode=True)
            break
        else:
            a=input("替换:")
            b=input("替换为:")
            rep[a]=b
            with open("replace.yaml", 'w', encoding="utf-8") as file:
                yaml.dump(rep, file, allow_unicode=True)
        print("----------------")

def main1():
    checkD()
    print("读取文件列表")
    la = os.listdir("docx")
    # 读取文档对象

    for i in la:
        print("========================")
        print("开始替换")
        print("当前文件:"+i)
        print("当前替换任务词典:" + str(rep))
        sa=input("按1进入关键词增加/删除\n按任意键开始替换\n输入：")
        if sa=="1":
            checkD()
        doc = docx.Document("docx/"+i)
        # 定义要替换的字符和替换后的字符
        for key in rep:
            old_text = key
            new_text = rep.get(key)
            # 遍历文档中的所有段落对象
            for paragraph in doc.paragraphs:
                # 如果段落中包含要替换的字符
                if old_text in paragraph.text:
                    # 获取段落中的所有运行对象
                    runs = paragraph.runs
                    # 遍历运行对象列表
                    for run in runs:
                        # 如果运行对象中包含要替换的字符
                        if old_text in run.text:
                            # 用替换后的字符替换运行对象中的文本
                            run.text = run.text.replace(old_text, new_text)
        # 保存文档
        print("替换完成，保存文档:"+i)
        doc.save("newDocx/"+i)

if __name__ == '__main__':
    print("检查目录是否存在.....")
    if os.path.exists("docx"):
        print("ok")
    else:
        os.mkdir("docx")
    if os.path.exists("newDocx"):
        print("ok")
    else:
        os.mkdir("newDocx")
    if os.path.exists("replace.yaml"):
        print("检查是否存在本地替换词库.....")
        print("ok")
    else:
        print("初始化替换词......")
        rep["芝士初始化替换关键词11111"]="该不会真能替换到东西吧"
        with open("replace.yaml", 'w', encoding="utf-8") as file:
            yaml.dump(rep, file, allow_unicode=True)

    with open("replace.yaml", 'r', encoding='utf-8') as f:
        result = yaml.load(f.read(), Loader=yaml.FullLoader)
    rep=result
    print("读取到本地替换字典:"+str(rep))
    print("欢迎使用本程序，请将源文件放在docx文件夹下")
    input("按任意键继续")
    print("////////////////////")
    main1()
    input("程序执行完成，按任意键退出")
