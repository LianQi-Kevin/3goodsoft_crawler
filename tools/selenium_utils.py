import logging
from typing import Literal

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


def set_driver(headless_mode: bool = False, auto_detach: bool = False,
               download_path: str = None, proxy: str = None) -> webdriver.Chrome:
    """
    Set up the driver
    :param proxy: The Proxy IP, like: http://127.0.0.1:10800
    :param download_path: if not None, change default download path
    :param auto_detach: whether to automatically detach the driver
    :param headless_mode: Whether to use headless mode
    """
    options = Options()
    # 无头模式
    if headless_mode:
        logging.info("Use headless mode")
        options.add_argument('headless')
        options.add_argument('blink-settings=imagesEnabled=false')  # 不加载图片, 提升速度

    # 进程结束自动关闭浏览器
    if not auto_detach:
        options.add_experimental_option("detach", True)

    # prefs基础设置
    prefs = {
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    }

    # 修改下载设置
    if download_path is not None:
        prefs["download.default_directory"] = download_path
    options.add_experimental_option("prefs", prefs)

    # 代理
    if proxy is not None:
        options.add_argument(f'--proxy-server={proxy}')

    # 常用设置
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-gpu')

    # 隐藏特征
    options.add_argument('ignore-certificate-errors')
    options.add_argument(
        'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/109.0.5414.74 Safari/537.36')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    options.page_load_strategy = "normal"
    driver = webdriver.Chrome(options=options, service=ChromeService(ChromeDriverManager().install()))
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """
            Object.defineProperty(navigator, 'webdriver', {
              get: () => undefined
            })
          """})

    return driver


def web_wait(driver: webdriver, by: Literal[
    "id", "xpath", "name", "class name", "tag name", "css selector", "link text", "partial link text"], element: str,
             until_sec: int = 20, wait_for_invisibility: bool = False, wait_for_clickable: bool = False):
    """
    等待一个元素出现、消失或变得可点击。

    :param driver: Selenium WebDriver实例。
    :param by: 定位元素的方法，如 By.ID 或 By.XPATH。
    :param element: 要定位的元素的标识符。
    :param until_sec: 最大等待时间（秒）。
    :param wait_for_invisibility: 如果为True，则等待元素消失；默认为False。
    :param wait_for_clickable: 如果为True，则等待元素变得可点击；默认为False。
    """
    if wait_for_clickable:
        WebDriverWait(driver, until_sec).until(EC.element_to_be_clickable((by, element)))
    elif wait_for_invisibility:
        WebDriverWait(driver, until_sec).until(EC.invisibility_of_element_located((by, element)))
    else:
        WebDriverWait(driver, until_sec).until(EC.presence_of_element_located((by, element)))


def check_element_exists(driver: webdriver, element: str, find_model=By.CLASS_NAME) -> bool:
    """
    Check whether the element exists
    :param driver: browser drive
    :param element: WebElement
    :param find_model: The selenium locator, default By.CLASS_NAME
    :return: bool, whether the element exists
    """
    try:
        driver.find_element(find_model, element)
        return True
    except Exception:
        return False
