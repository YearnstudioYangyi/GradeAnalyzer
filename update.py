import os
import requests # type: ignore
import sys
import shutil

def SendGetRequest(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # 如果响应状态码不是200，会抛出异常
        return response.json()  # 如果响应是JSON格式，可以解析为字典
    except requests.exceptions.RequestException as e:
        print(f"请求失败: {e}")
        return None

def download_file(url, destination):
    response = requests.get(url, stream=True)
    response.raise_for_status()
    with open(destination, 'wb') as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)

if __name__ == "__main__":
    js = SendGetRequest("https://db.yearnstudio.cn/static/cj/version.json")
    if js is None:
        print("请求失败，请检查网络连接")
        sys.exit(1)
    if js.get('version') != "1.0.0":
        print("发现新版本，准备下载")
        download_file(js.get('url'), "main.exe.new")
        # 删除main.exe
        os.remove("main.exe")
        # 重命名
        shutil.move("main.exe.new", "main.exe")