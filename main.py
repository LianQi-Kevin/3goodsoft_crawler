import csv
import logging
import time
from typing import Optional

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from pydantic import BaseModel, Field

from tools.logging_utils import log_set
from tools.selenium_utils import set_driver, web_wait
from tools.tools import flatten_dict, match_clean


MAIN_URL = "https://www.3goodsoft.net/campus/"
USERNAME = "Your Username && Password"


class QuestionOptionsModel(BaseModel):
    A: Optional[str] = Field(default=None)
    B: Optional[str] = Field(default=None)
    C: Optional[str] = Field(default=None)
    D: Optional[str] = Field(default=None)
    E: Optional[str] = Field(default=None)


class QuestionModel(BaseModel):
    index: str
    type: str
    description: str
    options: QuestionOptionsModel
    answer: str
    analysis: str


def login(driver: webdriver.Chrome, user_name: str):
    """Login to the website."""
    web_wait(driver, By.CSS_SELECTOR, "div.index_header button")
    ActionChains(driver).move_to_element(
        driver.find_element(By.CSS_SELECTOR, "div.index_header button")).click().perform()
    web_wait(driver, By.CSS_SELECTOR, "div.loginDialog div.input_box div.el-input")
    for index, element in enumerate(driver.find_elements(By.CSS_SELECTOR, "div.loginDialog div.input_box input")):
        time.sleep(0.5)
        ActionChains(driver).move_to_element(element).click().send_keys(user_name).perform()
    ActionChains(driver).move_to_element(
        driver.find_element(By.CSS_SELECTOR, "div.loginDialog div.input_box button")).click().perform()


def switch_to_sub_page(driver: webdriver.Chrome):
    """Switch to '职业技能中心' page."""
    web_wait(driver, By.XPATH, "//div[@class='index_left']//div[text()='职业技能中心']", wait_for_clickable=True)
    time.sleep(0.5)
    # todo: 可能无法切换到‘职业技能中心’菜单
    ActionChains(driver).move_to_element(driver.find_element(By.XPATH,
                                                             "//div[@class='index_left']//div[text()='职业技能中心']")).double_click().perform()
    time.sleep(0.5)
    for tab in driver.window_handles:
        driver.switch_to.window(tab)
        if driver.current_url.startswith("https://www.3goodsoft.net/skill/home"):
            break


def switch_to_sub_page2(driver: webdriver.Chrome):
    """Switch to '顺序练习' page."""
    # todo: 可能无法切换到‘中级’菜单
    time.sleep(0.5)
    web_wait(driver, By.XPATH, "//div[text()='中级']", wait_for_clickable=True)
    ActionChains(driver).move_to_element(driver.find_element(By.XPATH, "//div[text()='中级']")).click().perform()
    time.sleep(0.5)
    web_wait(driver, By.XPATH, "//span[text()='顺序练习']", wait_for_clickable=True)
    ActionChains(driver).move_to_element(driver.find_element(By.XPATH, "//span[text()='顺序练习']")).click().perform()

    # answer_sheet
    sheet_options_element = "//div[@class='mockAnswerQuestion']//div[@class='answer-sheet']//div[@class='sheet-options']"
    web_wait(driver, By.XPATH, sheet_options_element)
    for num_page_element in driver.find_element(By.XPATH, sheet_options_element).find_elements(By.CSS_SELECTOR, "div.opt-item"):
        # 点击题号
        ActionChains(driver).move_to_element(num_page_element).click().perform()
        web_wait(driver, By.XPATH, "//div[@class='mockAnswerQuestion']//div[@class='answer-sheet']//div[@class='sheet-options']")

        # 提取信息
        web_wait(driver, By.CSS_SELECTOR, "div.question-about")
        question_element = driver.find_element(By.CSS_SELECTOR, "div.question-about")
        # 展开答案信息
        web_wait(question_element, By.CSS_SELECTOR, "div.actions div.a-right button", wait_for_clickable=True)
        ActionChains(driver).move_to_element(driver.find_element(By.CSS_SELECTOR, "div.actions div.a-right button")).click().perform()
        web_wait(question_element, By.CSS_SELECTOR, "div.result span")
        try:
            question = QuestionModel(**{
                "index": match_clean(question_element.find_element(By.CSS_SELECTOR, "div.topic div.t-number span.idx").text),
                "type": match_clean(question_element.find_element(By.CSS_SELECTOR, "div.topic div.t-number span.el-tag__content").text),
                "description": match_clean(question_element.find_element(By.CSS_SELECTOR, "div.topic div.t-answer").text),
                "options": {
                    match_clean(option.find_element(By.CSS_SELECTOR, "div.t-key").text): match_clean(option.find_element(By.CSS_SELECTOR, "div.t-value").text)
                    for option in question_element.find_elements(By.CSS_SELECTOR, "div.topic div.t-option")
                },
                "answer": match_clean(question_element.find_element(By.CSS_SELECTOR, "div.result span.r-right").text),
                "analysis": match_clean(question_element.find_element(By.CSS_SELECTOR, "div.analysis div.a-des").text)
            })
            logging.info(question)
            yield question

        except Exception as e:
            logging.error(e)


def main():
    log_set(log_save=False, log_level=logging.INFO)
    driver = set_driver(headless_mode=False, auto_detach=True)
    driver.get(MAIN_URL)
    login(driver, USERNAME)
    switch_to_sub_page(driver)

    # 初始化CSV, 以dict模式写入
    with open("answer_sheet.csv", "w", encoding="UTF-8") as f:
        writer = csv.DictWriter(
            f=f,
            fieldnames=("index", "type", "description", "options_A", "options_B", "options_C", "options_D", "options_E",
                        "answer", "analysis"),
            lineterminator='\n',
        )
        writer.writeheader()

        for index, question in enumerate(switch_to_sub_page2(driver)):
            # 写入CSV
            data = flatten_dict(question.model_dump())
            # 写入CSV
            writer.writerow(data)

            # n次存储保存一次
            if index % 10 == 0:
                f.flush()


if __name__ == '__main__':
    main()
