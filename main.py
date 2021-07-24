import base64
import time
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from smtplib import SMTP_SSL

import execjs
import requests
import yaml
from aip import AipOcr
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

ROOT_DIR = Path(__file__).resolve(strict=True).parent

# 文件路径
imagePath = str(ROOT_DIR / "gkcx.png")  # image保存路径
htmlPath = str(ROOT_DIR / "gkcx.html")  # html保存路径
jsPath = str(ROOT_DIR / "lz-string.js")  # js路径
cfgPath = str(ROOT_DIR / "config.yaml")  # 配置文件路径


def initCfg() -> dict:
    """初始化配置"""
    with open(cfgPath) as f:
        data = yaml.safe_load(f)
    return data


# 加载配置
cfg = initCfg()

# 设置表头
headers = {
    'Origin': 'http://gkcf.jxedu.gov.cn',
    'Content-Type': 'application/x-www-form-urlencoded',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Referer': 'http://gkcf.jxedu.gov.cn/',
    'Accept-Language': 'zh-CN,zh;q=0.9'
}


def compressToBase64(key: str) -> str:
    """对字段进行处理"""
    return execjs.compile(open(jsPath).read()).call("LZString.compressToBase64", key)


def getCodeImg(session: requests.Session) -> bytes:
    """获取验证码"""
    resp = session.get(f"http://gkcf.jxedu.gov.cn/captcha/getcode?t={int(time.time() * 1000)}", headers=headers)
    img = resp.json()["Data"]["Img"]
    img_bytes = base64.b64decode(img)
    return img_bytes


def cleanCode(word: str) -> str:
    """对验证码进行清理"""
    return word.strip()[:4]


def ocrCode(img: bytes) -> (bool, str):
    """识别验证码"""
    APP_ID = cfg["baiduOcr"]["APP_ID"]
    API_KEY = cfg["baiduOcr"]["API_KEY"]
    SECRET_KEY = cfg["baiduOcr"]["SECRET_KEY"]
    aipOcr = AipOcr(APP_ID, API_KEY, SECRET_KEY)
    resp = aipOcr.basicAccurate(img)
    try:
        code = cleanCode(resp["words_result"][0]["words"])
    except Exception as e:
        print(f"错误: {e}")
        return False, ""
    return True, code


def login(session: requests.Session, code: str) -> (bool, str):
    """开始登录"""
    data = {
        "key1": compressToBase64(cfg["login"]["KEY1"]),
        "key2": compressToBase64(cfg["login"]["KEY2"]),
        "key3": compressToBase64(code),
    }
    resp = session.post("http://gkcf.jxedu.gov.cn/", headers=headers, data=data)
    if "验证码错误" in resp.text:
        return False, ""
    # 处理js和css文件，能让浏览器正确加载
    text = resp.text
    text = text.replace('src="lib', 'src="http://gkcf.jxedu.gov.cn/lib')
    text = text.replace('src="js', 'src="http://gkcf.jxedu.gov.cn/js')
    text = text.replace('href="lib', 'href="http://gkcf.jxedu.gov.cn/lib')
    text = text.replace('href="css', 'href="http://gkcf.jxedu.gov.cn/css')
    with open(htmlPath, "w") as fw:
        fw.write(text)
    return True, text


def screenshot():
    """截图"""
    chrome_options = Options()

    # 服务器上运行需加上
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    # 设置无头运行
    chrome_options.add_argument('--headless')

    # 初始化
    driver = webdriver.Chrome(options=chrome_options)

    # 打开本地html文件
    driver.get(f"file://{htmlPath}")

    # 设置宽高，要不然截屏不对
    width = driver.execute_script("return document.documentElement.scrollWidth")
    height = driver.execute_script("return document.documentElement.scrollHeight")
    driver.set_window_size(width, height)

    # 保存图片
    time.sleep(1)
    driver.save_screenshot(imagePath)
    driver.close()


def sendMail():
    """发送图片邮件"""
    msg = MIMEMultipart()

    # 设置图片
    with open(imagePath, "rb") as fr:
        img = MIMEImage(fr.read())
    img.add_header('Content-ID', 'dns_config')
    msg.attach(img)

    # 构建html邮件内容
    mail_content = """
    <html>
      <body>
        <img src="cid:dns_config", width="500", height="800">
      </body>
    </html>
    """
    text = MIMEText(mail_content, 'html', 'utf-8')
    msg.attach(text)

    # 邮件主题描述
    msg["Subject"] = "GKCX"
    with SMTP_SSL(host=cfg["email"]["HOST"], port=cfg["email"]["PORT"]) as smtp:
        # 登录发送邮件服务器
        smtp.login(user=cfg["email"]["USER"], password=cfg["email"]["PASSWORD"])
        # 实际发送、接收邮件配置
        smtp.sendmail(from_addr=cfg["email"]["USER"], to_addrs=cfg["email"]["TO"], msg=msg.as_string())


def main():
    for num in range(10):
        session = requests.Session()
        img = getCodeImg(session=session)
        isCorrect, code = ocrCode(img)
        if not isCorrect:
            print(f"[{num}]验证码识别失败，重试...")
            continue
        isCorrect, text = login(session, code)
        if isCorrect:
            if "暂无录取信息" in text:
                print("结果还未出来...")
                return
            print("验证码识别成功，开始截图...")
            screenshot()
            print("截图成功，开始发送邮件...")
            sendMail()
            return
        else:
            print(f"[{num}]验证码识别错误，重试...")
            time.sleep(3)
    print("重试次数结束...")


if __name__ == '__main__':
    main()
