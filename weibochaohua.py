# -*- coding: utf-8 -*-
import csv
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait


browser = webdriver.Chrome()
#browser.set_window_size(500,600)
wait = WebDriverWait(browser, 10)
browser.get('https://weibo.com/')

#登陆微博，不登陆只能请求10次
def login():
    username = '***'#用户名
    password = '***'#密码
    event = (By.XPATH, '//*[@id="pl_login_form"]')
    wait.until(EC.presence_of_element_located(event))
    browser.find_element_by_xpath('//*[@id="loginname"]').send_keys(username)
    time.sleep(1)
    browser.find_element_by_xpath('//*[@id="pl_login_form"]/div/div[3]/div[2]/div/input').send_keys(password)
    time.sleep(1)
    browser.find_element_by_xpath('//*[@id="pl_login_form"]/div/div[3]/div[6]/a').click()
    print('登陆成功')

def get_info(i,keyword):
    # #登陆前
    # input = wait.until(
    #     EC.presence_of_element_located((By.CSS_SELECTOR, '#plc_top > div > div > div.gn_search_v2 > input'))
    # )
    #登陆后                                                #pl_common_top > div > div > div.gn_search_v2 > input
    input = wait.until(
        EC.presence_of_element_located((By.XPATH, '//input[@class="W_input"]'))
     )
    #登陆前
    # submit = wait.until(EC.element_to_be_clickable(
    #     (By.CSS_SELECTOR, '#weibo_top_public > div > div > div.gn_search_v2 > a')))
    #
    # 登陆后
    submit = wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//a[@class="W_ficon ficon_search S_ficon"]')))


    input.clear()
    input.send_keys('#'+keyword+'#')

    submit.click()
    # 登陆前
    # link=browser.find_element_by_xpath('//*[@id="pl_weibo_direct"]/div/div[1]/div/div/div[2]/h1/a[@href]')
    # link.click()

    # 登陆后
    link=browser.find_element_by_xpath('//div[@class="detail"]/h1/a[@href]')
    link.click()
    browser.switch_to_window(browser.window_handles[i+1])
    wait.until(EC.element_to_be_clickable((By.XPATH,'//td[@class="S_line1"]/strong')))
    contents = browser.find_elements_by_xpath('//td[@class="S_line1"]/strong')
    item = {}
    try:
        item['名字'] = keyword
        item['阅读'] = contents[0].text
        item['帖子'] = contents[1].text
        item['粉丝'] = contents[2].text
        print(item)
        # save_csv('weibo.csv',item)
    except IndexError:
        pass

def save_csv(filename,data):
    with open(filename, 'a',newline='', errors='ignore') as f:
        writer =csv.writer(f)
        #writer.writerow(('名字','阅读','帖子','粉丝'))
        writer.writerow(data.values())



if __name__ == '__main__':

    import pandas as pd
    login()
    time.sleep(10)
    keywords = pd.read_excel('半年报数据-微博.xlsx', sheel='超话题')
    for i,keyword in enumerate(keywords['name']):
        print(i+1,keyword)
        get_info(i,keyword)
