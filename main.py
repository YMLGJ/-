import schedule
import time
import subprocess
import datetime

def run_another_script():
    subprocess.call(['python', 'Weibo_Trending_Topics_Extraction.pyw'])


def check():
    now = datetime.datetime.now()
    formatted_now = now.strftime("%H:%M")
    minute = now.minute
    print(f'{formatted_now} 分钟检测中......')
    if minute % 2 == 0:
        run_another_script()
        print('开始运行！')
    else:
        print('未到运行时间。')


# 使用schedule库来安排check_and_run函数每分钟运行一次
schedule.every(1).minute.do(check)
k=0
k=int(input('请输入数字“1”开始脚本：'))
if k==1:
    while True:
        schedule.run_pending()
        time.sleep(1)
