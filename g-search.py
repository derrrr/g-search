import re
import os
import sys
import time
import codecs
import random
import shutil
import requests
import configparser
import pandas as pd
from pathlib import Path
from chardet import detect
from datetime import datetime
from selenium import webdriver
from urllib.parse import unquote, urlparse
from bs4 import BeautifulSoup as BS
from selenium.webdriver.chrome.options import Options


def _load_config():
    config_path = "./config.ini"
    with open(config_path, "rb") as ef:
        config_encoding = detect(ef.read())["encoding"]
    config = configparser.ConfigParser()
    config.read_file(codecs.open(config_path, "r", config_encoding))
    return config


def _requests_retry_session(config, status_forcelist=(500, 502, 504), session=None):
    session = requests.session()
    headers = {"user-agent": config["Requests_header"]["user-agent"]}
    session.headers.update(headers)
    return session


def isRational(txt):
    try:
        float(txt)
        return True
    except ValueError:
        return False


class G_search:
    def __init__(self):
        self.config = _load_config()
        self.project_list = self.get_project()
        self.page_dict = self.Google_page()
        self.date_str = datetime.today().strftime("%Y%m%d")
        self.home_path = str(Path.home()).replace("\\", "/")
        self.chrome_opt = self.selenium_setting()

    def get_project(self):
        excel_dir = "./excel"
        if not os.path.exists(excel_dir):
            os.makedirs(excel_dir)
        project = [os.path.splitext(filename)[0] for filename in os.listdir(excel_dir)]
        print("Get project: {}\n".format(", ".join(project)))
        for p in project:
            if not os.path.exists("./project/{}".format(p)):
                os.makedirs("./project/{}".format(p))
        return project

    def get_keyword_and_target(self, project_file):
        project_dir = "./excel"
        project_path = "{}/{}.xlsx".format(project_dir, project_file)
        all_sheet = pd.ExcelFile(project_path).sheet_names
        attach_1 = [s for s in all_sheet if "附件一" in s][0]
        attach_2 = [s for s in all_sheet if "附件二" in s][0]

        ## Target
        pre_target = pd.read_excel(project_path, sheet_name=attach_1, skiprows=1)
        # Find the first row
        row_target = pre_target.loc[pre_target[pre_target.columns[0]] == 1].index.values.astype(int)[0]
        # Load dataframe
        target_cols = ["序號", "標題", "網址"]
        sheet_target = pd.read_excel(project_path, sheet_name=attach_1, skiprows=row_target + 1, usecols=range(3))
        sheet_target.rename(columns=dict(zip(sheet_target.columns, target_cols)), inplace=True)
        sheet_target = sheet_target[sheet_target["序號"].apply(lambda x: isRational(x))].reset_index(drop=True).dropna(subset=["序號"])
        sheet_target["序號"] = sheet_target["序號"].values.astype(int)

        op_dir = "./project/{}/operation".format(self.project_name)
        if not os.path.exists(op_dir):
            os.makedirs(op_dir)
        target_path = "{}/{}_{}_target.csv".format(op_dir, self.date_str, self.project_name)
        sheet_target[["序號", "標題", "網址"]].to_csv(target_path, index=False, encoding="utf-8-sig")

        ## Keyword
        pre_sheet = pd.read_excel(project_path, sheet_name=attach_2)
        # Find the first row
        row_header = pre_sheet.loc[pre_sheet[pre_sheet.columns[0]] == 1].index.values.astype(int)[0]
        # Load dataframe
        op_sheet = pd.read_excel(project_path, sheet_name=attach_2, skiprows=row_header).fillna(method="ffill")
        # Replace "\n" in headers
        op_sheet.columns = op_sheet.columns.str.replace("\n", "")
        # Change first column header to "W"
        op_sheet.columns.values[0] = "W"

        keyword_path = "{}/{}_{}_keyword.csv".format(op_dir, self.date_str, self.project_name)
        op_sheet[["W", "操作目標字"]].to_csv(keyword_path, index=False, encoding="utf-8-sig")

        return sheet_target[["序號", "標題", "網址"]].values.tolist(), \
                op_sheet[["W", "操作目標字"]].values.tolist()

    def Google_page(self):
        page_key = ["第一頁", "第二頁", "第三頁"]
        page_parameter = [0, 10, 20]
        return dict(zip(page_key, page_parameter))

    def html_preprocess(self, key_word, count):
        url = "http://www.google.com/search?q={}&hl=zh-TW&ie=utf-8&oe=utf-8&start={}".format(key_word, count)

        # requests was banned
        # res = self.rs.get(url, timeout=9)
        # res_text = res.text
        # soup = BS(res_text, "lxml")

        driver = webdriver.Chrome(
            executable_path=self.config["Chrome_Canary"]["CHROMEDRIVER_PATH"],
            options=self.chrome_opt
        )
        driver.get(url)
        res_text = driver.page_source
        soup = BS(res_text, "lxml")
        driver.quit()

        # Set "utf-8"
        soup.find("meta")["charset"] = "utf-8"

        # debug
        # debug_soup = soup.prettify()
        # debug_html_dir = "./project/{}/debug".format(self.project_name)
        # if not os.path.exists(debug_html_dir):
        #     os.makedirs(debug_html_dir)
        # with open("{}/res_debug_{}.html".format(debug_html_dir, self.date_str), "w", encoding="utf-8") as save:
        #     save.write(debug_soup)

        if soup.find(id="recaptcha"):
            recapt_continue = soup.find("input", {"name": "continue"})["value"]
            recapt_q = soup.find("input", {"name": "q"})["value"]
            recapt_url = "https://www.google.com/sorry/index?continue={}&q={}".format(recapt_continue, recapt_q)
            print(recapt_url)
            print("被Google ban了QQ")
            print("請換IP或手動解reCAPTCHA(手解不一定有效)或等到Google自己解除\n")
            sys.exit()

        # Set footcnt as visible
        soup.find(id="footcnt")["style"] = "position:relative;visibility:visible"

        prettified = soup.prettify()

        # Save origin html with utf-8 encoding
        origin_html_dir = "./project/{}/origin".format(self.project_name)
        if not os.path.exists(origin_html_dir):
            os.makedirs(origin_html_dir)
        with open("{}/res_origin_{}.html".format(origin_html_dir, self.date_str), "w", encoding="utf-8") as save:
            save.write(prettified)

        # Replace url with prefix from original src
        src_sub = {"//ssl.gstatic.com": "http://ssl.gstatic.com", \
                   "/images/nav_logo242.png": "http://www.google.com.tw/images/nav_logo242.png"}
        src_sub = dict((re.escape(k), v) for k, v in src_sub.items())
        pattern = re.compile("|".join(src_sub.keys()))
        prettified = pattern.sub(lambda m: src_sub[re.escape(m.group(0))], prettified)

        soup_p = BS(prettified, "lxml")
        # Drop ads on top and bottom
        for ad in soup_p.find_all(class_="C4eCVc"):
            ad.decompose()
        # Drop ads on right column
        for right_col in soup_p.find_all(class_="cu-container"):
            right_col.decompose()
        # Remove the privacy check
        for check in soup_p.find_all(id="taw"):
            check.decompose()
        # Remove the privacy options
        for check in soup_p.find_all(role="dialog"):
            check.decompose()
        # Remove Chrome version check
        for version_check in soup_p.find_all(class_="gb_Fd gb_Zc"):
            version_check.decompose()
        # Add Google map image src
        for g_map in soup_p.find_all(id="lu_map"):
            g_map["src"] = "http://www.google.com.tw{}".format(g_map["src"])
        # Add url prefix to img src
        add_prefix = lambda src: "http://www.google.com.tw{}".format(src)
        src_to_fix = [soup_p.find(itemprop="image")["content"], \
                      soup_p.find(class_="logo").find("a").find("img")["src"]]
        src_fixed = list(map(add_prefix, src_to_fix))
        soup_p.find(itemprop="image")["content"], \
            soup_p.find(class_="logo").find("a").find("img")["src"] = src_fixed
        # Remove the background color of Google Apps
        style = soup.find("style", text=re.compile("gb_Dd")).text
        style_fix = soup.find("style", text=re.compile("gb_Dd")).string.replace(";background-color:#4d90fe", "")
        soup.find("style", text=re.compile("gb_Dd")).string = style_fix
        # Save no-ads html
        no_ads_dir = "./project/{}/no_ads".format(self.project_name)
        if not os.path.exists(no_ads_dir):
            os.makedirs(no_ads_dir)
        with open("{}/no-ads_{}.html".format(no_ads_dir, self.date_str), "w", encoding="utf-8") as save:
            save.write(soup_p.prettify())
        return soup_p.prettify()

    def process_check(self):
        result_dir = "./project/{}/result".format(self.project_name)
        if not os.path.exists(result_dir):
            os.makedirs(result_dir)
        result_path = "{}/result_{}_{}.csv".format(result_dir, self.project_name, self.date_str)
        if os.path.exists(result_path):
            with open(result_path, "r", encoding="utf-8-sig") as check:
                if check.read()[:3] == "序號,":
                    return None, None
            with open(result_path, "r", encoding="utf-8-sig") as check:
                pairs = [line.strip().split(",") for line in check.readlines()]
                url_index = [pair[0] for pair in pairs]
                keyword_index = [pair[1] for pair in pairs]
                return int(keyword_index[-1]) - 1, int(url_index[-1]),
        else:
            return 0, 0

    def search_html(self, html, key_word, page_x, page_count):
        self.found = 0
        soup_no_ad = BS(html, "lxml")
        title_slice = int(self.config["Title_part"]["slice"])
        url_pat = re.compile("(\/([\d\-\/]*))-.*")
        for target in self.target_list[self.url_last:]:
            rank = 1
            self.search = 0
            for s_res in soup_no_ad.find(id="rso").find_all(class_="g"):
                title_part = re.sub("\s*\.*\\n\s*", "", str(s_res.a.string))
                if unquote(target[2]) in unquote(s_res.a["href"]) and url_pat.sub("\\1", unquote(target[2])) == url_pat.sub("\\1", unquote(s_res.a["href"])) \
                    or (target[1][:title_slice] == title_part[:title_slice] and urlparse(unquote(target[2])).hostname == urlparse(unquote(s_res.a["href"])).hostname):
                    s_res.find(class_="rc")["style"] = "border-width:2px; border-style:solid; border-color:red; padding:1px;"
                    message = "關鍵字: {} {}\t在 第{}頁 第{}個 找到\n{}".format(\
                        key_word[0], key_word[1], page_count, rank, target[2])
                    print(message)
                    self.found = 1
                    self.search = 1
                    result_row = "{},{},\"{}\",\"{}\",\"{}\",{}, {}\n".format(\
                        target[0], key_word[0], key_word[1], target[1], target[2], page_x, 1)
                    self.search_result(result_row)
                rank += 1
            if self.search == 0:
                result_row = "{},{},\"not found and will be remove\",,,,\n".format(target[0], key_word[0])
                self.search_result(result_row)
        self.url_last = 0

        frame_dir = "./project/{}/frame/{}".format(self.project_name, self.date_str)
        if self.found == 1:
            if not os.path.exists(frame_dir):
                os.makedirs(frame_dir)
            save_path = "{}/W{}_{}_{}_P{}.html".format(frame_dir, key_word[0], key_word[1], self.date_str, page_count)
            with open(save_path, "w", encoding="utf-8") as save:
                save.write(soup_no_ad.prettify())
            return os.path.abspath(save_path).replace("\\", "/")

    def search_result(self, result_content):
        result_dir = "./project/{}/result".format(self.project_name)
        if not os.path.exists(result_dir):
            os.makedirs(result_dir)
        result_path = "{}/result_{}_{}.csv".format(result_dir, self.project_name, self.date_str)
        with open(result_path, "a", encoding="utf-8-sig") as result:
            result.write("{}".format(result_content))

    def result_end(self):
        result_dir = "./project/{}/result".format(self.project_name)
        result_path = "{}/result_{}_{}.csv".format(result_dir, self.project_name, self.date_str)
        # if os.path.exists(result_path):
        df = pd.read_csv(result_path, encoding="utf-8-sig", header=None, engine="python")
        res_cols = ["序號", "W", "操作關鍵字", "標題", "操作網址", "搜尋結果頁", datetime.today().strftime("%Y/%m/%d")]
        df.columns = res_cols
        # Clear the not found lines
        df = df[df["操作關鍵字"].str.contains("not found and will be remove") == False]
        # Drop duplicates
        df.drop_duplicates(keep="first", inplace=True)
        # Sort the result
        df["page"] = df["搜尋結果頁"].map(self.page_dict)
        df = df.sort_values(["W", "序號", "page"], ascending=[True, True, True])
        # Drop rows as a keyword and a target found more than 1 page
        # df = df[~df[["序號", "W"]].duplicated(keep="first")]
        df = df.drop(labels=["page"], axis=1)
        df.to_csv(result_path, index=False, encoding="utf-8-sig")

    def concat(self):
        result_dir = "./project/{}/result".format(self.project_name)
        csv_files = []
        for dirpath, subdirs, files in os.walk(result_dir):
            for x in files:
                if x.endswith(".csv"):
                    csv_files.append(os.path.join(dirpath, x).replace("\\", "/"))
        dfs = [pd.read_csv(f, encoding="utf-8-sig", engine="python") for f in csv_files]
        df = pd.concat(dfs, sort=False, ignore_index=True)
        df = df.groupby(["序號", "W", "操作關鍵字", "標題", "操作網址", "搜尋結果頁"]).sum().reset_index()
        df.iloc[:, 6:] = df.iloc[:, 6:].values.astype(int)
        # Sort the concatenated dataframe
        df["page"] = df["搜尋結果頁"].map(self.page_dict)
        df = df.sort_values(["W", "序號", "page"], ascending=[True, True, True])
        df = df.drop(labels=["page"], axis=1)
        concat_dir = "./project/{}/concat".format(self.project_name)
        if not os.path.exists(concat_dir):
            os.makedirs(concat_dir)
        concat_path = "{}/concat_{}_{}.csv".format(concat_dir, self.project_name, self.date_str)
        df.to_csv(concat_path, index=False, encoding="utf-8-sig")

    def remove_temp_dir(self):
        dir_list = ["no_ads", "origin"]
        for rm_dir in dir_list:
            rm_path = "./project/{}/{}".format(self.project_name, rm_dir)
            if os.path.exists(rm_path):
                shutil.rmtree(rm_path)

    def check_screenshot(self):
        frame_dir = "./project/{}/frame/{}".format(self.project_name, self.date_str)
        screenshot_dir = "./project/{}/screenshot/{}".format(self.project_name, self.date_str)
        frame_num = len(os.listdir(frame_dir))
        shot_num = len(os.listdir(screenshot_dir))
        if frame_num > shot_num:
            print("html數量和screen數量不一致 html: {} screenshot: {}".format(frame_num, shot_num))
            print("{} 重截圖".format(self.project_name))
            self.re_screenshot()
            print("{} 重截圖完成\n".format(self.project_name))
        elif frame_num < shot_num:
            print("html數量和screen數量不一致 html: {} screenshot: {}".format(frame_num, shot_num))
            print("截圖比html多！？")

    def re_screenshot(self):
        frame_dir = "./project/{}/frame/{}".format(self.project_name, self.date_str)
        screenshot_dir = "./project/{}/screenshot/{}".format(self.project_name, self.date_str)
        frame_files = []
        for dirpath, subdirs, files in os.walk(frame_dir):
            for x in files:
                if x.endswith(".html"):
                    frame_files.append([os.path.join(dirpath, x).replace("\\", "/"), x[:-5]])

        screenshot_files = []
        for dirpath, subdirs, files in os.walk(screenshot_dir):
            for x in files:
                if x.endswith(".png"):
                    screenshot_files.append(x[:-4])

        for file in frame_files:
            if not file[1] in screenshot_files:
                driver = webdriver.Chrome(
                    executable_path=self.config["Chrome_Canary"]["CHROMEDRIVER_PATH"],
                    options=self.chrome_opt
                )
                driver.get("file:///{}".format(file[0]))
                width  = driver.execute_script("return Math.max(document.body.scrollWidth, document.body.offsetWidth, document.documentElement.clientWidth, document.documentElement.scrollWidth, document.documentElement.offsetWidth);")
                height = driver.execute_script("return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
                driver.set_window_size(width, height)
                print("screenshot: {}.png".format(file[1]))
                save_path = "{}/{}.png".format(screenshot_dir, file[1])
                driver.save_screenshot(save_path)
                driver.quit()

    def selenium_setting(self):
        # Selenium setting
        chrome_opt = Options()
        # chrome_opt.set_headless(headless=True)
        chrome_opt.headless = True
        chrome_opt.add_argument("--disable-gpu")
        # chrome_opt.add_argument("--disable-software-rasterizer")
        chrome_opt.add_argument("--mute-audio")
        # chrome_opt.add_argument("--remote-debugging-port=9222")
        chrome_opt.add_argument("--ignore-gpu-blocklist")
        chrome_opt.add_argument("--no-default-browser-check")
        # chrome_opt.add_argument("--no-first-run")
        chrome_opt.add_argument("--disable-default-apps")
        # chrome_opt.add_argument("--disable-infobars")
        chrome_opt.add_argument("--disable-extensions")
        # chrome_opt.add_argument("--test-type")
        chrome_opt.add_argument("--hide-scrollbars")
        chrome_opt.add_experimental_option("excludeSwitches", ["enable-logging"])
        chrome_opt.binary_location = self.config["Chrome_Canary"]["CHROME_PATH"].format(self.home_path)
        return chrome_opt

    def screenshot(self, html_path, key_word, page_count):
        screenshot_dir = "./project/{}/screenshot/{}".format(self.project_name, self.date_str)
        if not os.path.exists(screenshot_dir):
            os.makedirs(screenshot_dir)

        driver = webdriver.Chrome(
            executable_path=self.config["Chrome_Canary"]["CHROMEDRIVER_PATH"],
            options=self.chrome_opt
        )
        driver.get("file:///{}".format(html_path))
        width  = driver.execute_script("return Math.max(document.body.scrollWidth, document.body.offsetWidth, document.documentElement.clientWidth, document.documentElement.scrollWidth, document.documentElement.offsetWidth);")
        height = driver.execute_script("return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
        driver.set_window_size(width, height)
        save_path = "{}/W{}_{}_{}_P{}.png".format(screenshot_dir, key_word[0], key_word[1], self.date_str, page_count)
        driver.save_screenshot(save_path)
        driver.quit()

    def process(self):
        start_time = datetime.now().replace(microsecond=0)

        min_sleep = float(self.config["Sleep_time"]["min"])
        max_sleep = float(self.config["Sleep_time"]["max"])

        for project in self.project_list:
            self.project_name = project
            print("--{} 執行--".format(self.project_name))
            self.keyword_last, self.url_last = self.process_check()
            self.target_list, self.keyword_list = self.get_keyword_and_target(self.project_name)
            if self.keyword_last == None:
                print("--{} 已完成--\n".format(self.project_name))
                self.concat()
                self.remove_temp_dir()
                self.check_screenshot()
                continue
            elif self.keyword_last > 0 or self.url_last > 0:
                print("從第{}個關鍵字 第{}個目標網址 繼續\n".format(self.keyword_last + 1, self.url_last))
            keyword_count = 1
            for keyword in self.keyword_list[self.keyword_last:]:
                for page_key, page_parameter in self.page_dict.items():
                    self.rs = _requests_retry_session(self.config)
                    no_ad_html = self.html_preprocess(keyword[1], page_parameter)
                    if no_ad_html:
                        page_count = int(page_parameter / 10 + 1)
                        frame_path = self.search_html(no_ad_html, keyword, page_key, page_count)
                        if frame_path:
                            self.screenshot(frame_path, keyword, page_count)
                    sleep_time = random.uniform(min_sleep, max_sleep)
                    print("Sleep for {:.1f} secs.".format(sleep_time))
                    time.sleep(sleep_time)
                print("第{} / {}個關鍵字完成\t進度: {:.2%}\n".format(keyword_count + self.keyword_last, len(self.keyword_list), keyword[0]/len(self.keyword_list)))
                keyword_count += 1
            self.result_end()
            self.concat()
            self.remove_temp_dir()
            self.check_screenshot()
            print("--{} 已完成--\n".format(self.project_name))
        print("==全部完成 花費時間: {}==".format(str(datetime.now().replace(microsecond=0) - start_time)))

Gs = G_search()
Gs.process()