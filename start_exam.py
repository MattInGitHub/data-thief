import t_config
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from pyquery import PyQuery as pq
from openpyxl import Workbook
from time import sleep
import re
import subprocess

def down_img(index,url):
    name = str(index)+'_'+str(t_config.IMG_INDEX)
    file = './IMG/'+name+'.png'
    p = subprocess.Popen(["wget", "-O", file, "-q", "-t", "2", "-w", "1", "-c", url])
    p.wait()
    if (p.returncode == 1):
        return "[Insert:%s]" % url,''
    t_config.IMG_INDEX = t_config.IMG_INDEX + 1
    return "[Insert:%s]"%name,'=HYPERLINK("%s","%s")'%(file,name)

def replace_text(index,text,q_list,i_list):
    p_u = re.compile('<img.*?src=".*?"/>')
    p_l = re.compile('.*src="(.*?)"')
    key_list = p_u.findall(text)
    for key in key_list:
        match_l = p_l.match(key)
        if match_l:
            r_word,path = down_img(index, match_l.group(1))
            i_list.append(path)
            text = text.replace(key,r_word)
    q_list.append(text)


def login(driver):
    driver.get(t_config.LOGIN_URL)
    try:
        WebDriverWait(driver,t_config.TIMEOUT).until(EC.presence_of_element_located((By.CSS_SELECTOR,"input#loginphone")))
    except TimeoutException:
        print("超时")
    try:
        login_id = driver.find_element_by_css_selector("input#loginphone")
        sleep(0.5)
        login_id.send_keys(t_config.LOGIN_ID)
        sleep(0.5)
    except NoSuchElementException:
        print("没有找到ID框")
    try:
        login_pass = driver.find_element_by_css_selector("input#loginpassword")
        sleep(0.5)
        login_pass.send_keys(t_config.LOGIN_PASS)
        sleep(0.5)
    except NoSuchElementException:
        print("没有找到密码框")
    try:
        login_ok = driver.find_element_by_css_selector("#loginPasswordPage div.login.text-center")
        login_ok.click()
    except NoSuchElementException:
        print("没有找到提交")
    sleep(1)
    driver.get(t_config.LOGIN_URL)


def choose_cat(driver):
    try:
        WebDriverWait(driver, t_config.TIMEOUT).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.btn.btn-primary.select-paper")))
    except TimeoutException:
        print("超时")
    true_exam = driver.find_element_by_css_selector("span.btn.btn-primary.select-paper")
    true_exam.click()


def get_papers(driver):
    try:
        WebDriverWait(driver, t_config.TIMEOUT).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".all-paper")))
    except TimeoutException:
        print("超时")

    try:
        list = driver.find_elements_by_css_selector("div.paper-list div.paper")
        return list
    except NoSuchElementException:
        print("没有找到试卷 ")
        return None


def get_nav(driver):
    try:
        WebDriverWait(driver, t_config.TIMEOUT).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.chapter-nav")))
    except TimeoutException:
        print("超时")

    try:
        list = driver.find_elements_by_css_selector("div.box-inner > div.exercise-hd > div.fixed-head > div.chapter-nav > ul > li")
        return list
    except NoSuchElementException:
        print("没有找到导航 ")
        return None


def get_ques(driver,worksheet,cat):
    worksheet.append([cat])
    try:
        WebDriverWait(driver, t_config.TIMEOUT).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#question"+str(t_config.TOP_INDEX))))
    except TimeoutException:
        print("载入题目%d超时"%t_config.TOP_INDEX)
        exit(1)
    p_o = re.compile(r'^<div.*</span>(.*)</div>$')
    html = driver.page_source
    # 必须去掉xmlns属性，不然不能正确解析
    html = html.replace('xmlns', 'another_attr')
    doc = pq(html)
    for index in range(t_config.TOP_INDEX, t_config.TOP_INDEX + t_config.Q_COUNT):
        q_list = []
        i_list=[]
        all = doc("#question"+str(index))
        q = all('.overflow>p')
        q = q.__str__().replace('<p>','').replace('</p>','\n').strip('\n')
        t_config.IMG_INDEX=0
        replace_text(index, q,q_list,i_list)
        for i in range(1, 5):
            o = all('div.options > div:nth-child(%d)' % i)
            match_o = p_o.match(o.__str__())
            if match_o:
                o = match_o.group(1)
                replace_text(index, o, q_list, i_list)
        # print(q_list,'\n',i_list)
        if i_list:q_list.extend(i_list)
        worksheet.append(q_list)


if __name__ =='__main__':
    driver = t_config.DRIVER
    login(driver)
    choose_cat(driver)
    paper_list = get_papers(driver)
    btn_in = paper_list[0].find_element_by_css_selector("div.pull-right.button-wrap > span")
    title = paper_list[0].find_element_by_css_selector("div.name").text
    print(title)
    btn_in.click()

    wb = Workbook()
    dest_filename = 'real_exam.xlsx'
    ws = wb.create_sheet(title)

    pattern = re.compile(r'.*\[\d+/(\d+)\]')
    nav_list = get_nav(driver)
    top_index=0
    for nav in nav_list:
        nav.click()
        # print(nav.text)
        match = pattern.match(nav.text)
        if match:
            t_config.Q_COUNT = int(match.group(1))
            get_ques(driver,ws,nav.text)
            t_config.TOP_INDEX = t_config.TOP_INDEX+t_config.Q_COUNT
        sleep(2)
    wb.save(dest_filename)
