import time
import random
import logging
from threading import Thread
from threading import Semaphore
from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.common.exceptions import SessionNotCreatedException
from configparser import ConfigParser

CONF_FILE = "bo.conf"
LOG_FILE = "bo.log"
default_url = "https://www.2345.com/"

def init_log_config():
    logging.basicConfig(level=logging.DEBUG,format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',  \
        datefmt='%a, %d %b %Y %H:%M:%S',filename=LOG_FILE,filemode='a')

def get_random_idstring(min, max):
    id = random.randint(min, max)
    return str(id)

class ConfigManager():
    def __init__(self):
        self.__cp = ConfigParser()
        self.__cp.read(CONF_FILE)

    def get_skws(self):
        option_ls = self.__cp.options('search_key_word')
        skw_ls = list()
        for option in option_ls:
            skw_ls.append(self.__cp.get('search_key_word', option))
    
        logging.debug("Reading search key words : %s" % str(skw_ls))
        return skw_ls

    def get_random_skw(self):
        skw_ls = self.get_skws()
        index = random.randint(0, len(skw_ls)-1)
        skw = skw_ls[index]
        logging.debug("Getting key word : %s" % skw)
        return skw

    def get_sleep_sec(self):
        return self.__cp.getint('default', 'sleep_seconds')

    def get_max_windou(self):
        return self.__cp.getint('default', 'max_windou')

class Driver():
    def __init__(self):
        self.__cm = ConfigManager()
        self.__bd_num = 0
        self.__sem = Semaphore(1)
        self.__bd_max = self.__cm.get_max_windou()
        self.__sleep_sec = self.__cm.get_sleep_sec()

    def __get_one_broswer_driver(self):
        logging.debug("Creating new broswer driver...")
        self.__bd_num += self.__bd_num
        user_options = Options()
        user_options.add_argument('disable-infobars')

        try:
            bd = webdriver.Chrome(chrome_options=user_options)
        except SessionNotCreatedException as e:
            logging.error("Failed to create new broswer driver %s." % e.msg)
            return None
        
        logging.info("Created active broswer driver %d" % self.__bd_num)
        return bd
        
    def once_search_task(self):
        bd = self.__get_one_broswer_driver()
        if bd == None:
            return
        #bd.get(default_url)

        input_kw = bd.find_element_by_class_name("sch_inbox")
        input_kw = input_kw.find_element_by_name("word")
        input_kw.send_keys(self.__cm.get_random_skw())

        button = bd.find_element_by_id('j_search_sbm')
        button.click()
        time.sleep(2)

        bd.switch_to.window(bd.window_handles[1])
        #bd.refresh()

        re = bd.find_element_by_id("content_left")
        re = re.find_element_by_id(get_random_idstring(0, 10))
        re = re.find_element_by_class_name("c-showurl")
        re.click()

        bd.minimize_window()
        self.__sem.release()

        time.sleep(self.__sleep_sec)  
        self.__bd_num -= 1      
        bd.quit()

    def run(self):
        while (True):
            self.__sem.acquire()
            if (self.__bd_num == self.__bd_max):
                time.sleep(1)
                continue 
            
            th = Thread(target=self.once_search_task)
            th.setDaemon(True)
            th.start()
        

if __name__ == "__main__":
    init_log_config()
    driver = Driver()
    driver.run()
