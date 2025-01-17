import matplotlib # type: ignore
import pandas as pd # type: ignore
import numpy as np # type: ignore
import warnings
import os
import requests # type: ignore

# 屏蔽因为智学网导出造成的警告
warnings.filterwarnings('ignore')

# 配置matplotlib使用中文字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题


def rean_xlsx(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)
    # 将数据框转换为NumPy数组
    data_array = df.to_numpy()
    # 统计第一列的准考证号
    id = []
    for i in range(1,len(data_array)):
        id.append(data_array[i][0])
    # 获取第二列的班级
    _class = data_array[1][1]
    # 获取第三列的姓名
    name = []
    for i in range(1,len(data_array)):
        name.append(data_array[i][2])
    # 扫描第一行的科目
    subject = []
    for i in range(3,len(data_array[0])):
        if data_array[0][i] != '校次' and data_array[0][i] != '班次':
            subject.append(data_array[0][i])
    print("正在分析考试："+file_path.split('/')[-1])
    print("共" + str(len(subject) - 1) + "科目")

    # 创建数据JSON
    data = {"subject": subject,"class":_class,"student":{},"exam":file_path.split('/')[-1]}

    # 扫描学生数据
    for i in range(1,len(data_array)):
        # 创建学生JSON
        student = {"id":id[i-1],"name":name[i-1],"score":{}}
        # 扫描学生成绩及排名
        for j in range(3,len(data_array[0])):
            if data_array[0][j] != '校次' and data_array[0][j] != '班次':
                student["score"][data_array[0][j]] = data_array[i][j]
                for n in range(j+1,j+3):
                    student["score"][data_array[0][j] + data_array[0][n]] = data_array[i][n]
        data["student"][name[i-1]] = student
    print("共" + str(len(data["student"]) - 1) + "名学生")
    return data
    
def draw_pic(datas,student,subject='总分'):
    import matplotlib.pyplot as plt # type: ignore
    for i in range(len(datas)):
        if not (subject in datas[i]['subject'] and student in datas[i]['student']):
            print("数据错误,缺少学生或科目信息")
            return
    # 统计指定学生的每次考试成绩
    exam = []
    for i in range(len(datas)):
        exam.append(datas[i]['student'][student]['score'][subject])

    #提取考试名
    name = [data['exam'] for data in datas]
    # 绘制折线图
    plt.figure(figsize=(10, 6))
    plt.plot(name, exam, marker='o', linestyle='-', color='skyblue')
    plt.xlabel('考试')
    plt.ylabel(subject)
    plt.title(student + "的" + subject + "成绩分析图")
    if subject == '总分':
        plt.ylim(0, 800)  # 设置y轴范围
    else:
        plt.ylim(0, 150)
    plt.grid(True)  # 添加网格
    # 添加数据标签
    for i, txt in enumerate(exam):
        plt.annotate(txt, (name[i], exam[i]), textcoords="offset points", xytext=(0,10), ha='center')
    plt.show()

def send_request(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # 如果响应状态码不是200，会抛出异常
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"请求失败: {e}")
        return None

if __name__ == '__main__':
    version = "1.0.0"
    print("欢迎使用成绩分析工具，版本1.0.0\n")
    print("请将所有需要分析的考试数据放在data文件夹里")
    print("正在检查更新......")
    # 网络请求
    js = send_request("https://db.yearnstudio.cn/static/cj/version.json")
    try:
        if js.get('version') != version:
            print("发现新版本，请前往alist.yearnstudio.cn下载最新版本")
        else:
            print("当前版本为最新版本\n")
    except:
        print("网络错误，请稍后再试\n")
    if os.path.exists('./data') == False:
        print("data文件夹不存在，已创建")
        os.mkdir('./data')
    pause = input("在放好文件后，按任意键继续......\n")
    #列出data目录的所有xlsx文件
    files = os.listdir('./data')
    ar = []
    print("正在扫描xlsx文件\n")
    for i in range(len(files)):
        if files[i].endswith('.xlsx'):
            ar.append(rean_xlsx('./data/'+files[i]))
    print("扫描完成，共" + str(len(ar)) + "个xlsx文件\n")
    if len(ar) == 0 or len(ar) == 1:
        print("请至少输入两个xlsx文件")
        exit()
    print("考试数据顺序：")
    for i in range(len(ar)):
        print(str(i + 1) + ":" + ar[i]['exam'])
    print("是否需要更改顺序？是，请输入Y，否，请输入N")
    if input() == 'Y':
        while True:
            print("请输入需要更改的考试序号,输入0结束")
            num = input()
            if num == '0':
                break
            tager = input("请输入需要更改到的序号\n")
            ar.insert(int(tager) - 1, ar.pop(int(num) - 1))
    while True:
        student = input("请输入需要分析学生,输入q退出程序：")
        if student == 'q':
            exit()
        while True:
            subject = input("请输入需要分析科目,输入q重新选择学生：")
            if subject == 'q':
                break
            draw_pic(ar,student,subject)
