import os
import time
import random
import logging
import win32con
import win32api
import win32gui
import ctypes
import ctypes.wintypes
from threading import Thread
from threading import Semaphore
from threading import Lock
from threading import Event
from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import SessionNotCreatedException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from configparser import ConfigParser

CONF_FILE = "bo.conf"
LOG_FILE = "bo.log"
g_mutex = Lock()
g_sem = Semaphore(1)
g_run_event = Event()
g_run_event.set()

link_list = ('sina', 'sohu', 'ifeng', 'baidu', 'xinhuanet', 'weibo')

def init_log_config():
    if os.path.exists(LOG_FILE):
        fd = open(LOG_FILE, "rb+")
        fd.truncate()
        fd.close()
    cm = ConfigManager()
    logging.basicConfig(level=(cm.get_log_level() * 10),format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',  \
        datefmt='%a, %d %b %Y %H:%M:%S',filename=LOG_FILE,filemode='a')

def get_random_idstring(min, max):
    id = random.randint(min, max)
    return str(id)

class ConfigManager():
    def __init__(self):
        self.__cp = ConfigParser()
        self.__cp.read(CONF_FILE, encoding='utf-8-sig')
        self.__skw_keys = self.get_skws()

    def get_skws(self):
        option_ls = self.__cp.options('search_key_word')
        skw_ls = list()
        for option in option_ls:
            skw_ls.append(self.__cp.get('search_key_word', option))
    
        logging.info("Reading search key words : %s" % str(skw_ls))
        return skw_ls

    def get_random_skw(self):
        index = random.randint(0, len(self.__skw_keys)-1)
        skw = self.__skw_keys[index]
        logging.debug("Getting key word : %s" % skw)
        return skw

    def get_sleep_sec(self):
        return self.__cp.getint('default', 'sleep_seconds')

    def get_max_windou(self):
        return self.__cp.getint('default', 'max_windou')

    def get_default_url(self):
        return self.__cp.get('default', 'default_url')

    def get_log_level(self):
        return self.__cp.getint('default', 'log_level')

    def get_broswer_location(self):
        return self.__cp.get('default', 'broswer_location')

class Driver():
    def __init__(self):
        self.__cm = ConfigManager()
        self.__bd_num = 0
        self.__sleep_sec = self.__cm.get_sleep_sec()

    def __get_one_broswer_driver(self):
        logging.info("Creating new broswer driver...")
        print("Creating new broswer driver...")
        user_options = Options()
        user_options.add_argument('disable-infobars')
        #user_options.add_argument('--start-maximized')
        user_options.add_argument('--no-sandbox')
        user_options.add_argument('--disable-dev-shm-usage')
        user_options.add_argument('--incognito')

        user_options.binary_location = self.__cm.get_broswer_location()

        bd = None
        try:
            bd = webdriver.Chrome(chrome_options=user_options)
        except SessionNotCreatedException as e:
            logging.error("Failed to create new broswer driver, [SessionNotCreatedException] %s." % e.msg)
            print("Failed to create new broswer driver, [SessionNotCreatedException] %s." % e.msg)
        finally:
            g_mutex.acquire()
            self.__bd_num = self.__bd_num + 1
            g_mutex.release()
            logging.info("Created active broswer driver %d" % self.__bd_num)
            print("Created active broswer driver %d" % self.__bd_num)
            return bd
        
    def once_search_task(self):
        bd = self.__get_one_broswer_driver()
        if bd == None:
            g_sem.release()
            return
        bd.get(self.__cm.get_default_url())
        time.sleep(1)
        try:
            input_kw = bd.find_element_by_class_name("sch_inbox")
            input_kw = input_kw.find_element_by_name("word")
            input_kw.send_keys(self.__cm.get_random_skw())

            button = bd.find_element_by_id('j_search_sbm')
            button.click()
            time.sleep(2)

            t = 2
            while(t > 0):
                bd.switch_to.window(bd.window_handles[0])
                index = random.randint(0, 2)
                if index == 0:
                    bd.find_element_by_id("nav").find_element_by_class_name('txt').click()
                elif index == 1:
                    bd.find_element_by_class_name("cont-list").find_element_by_name('2').click()
                elif index == 2:
                    bd.find_element_by_class_name("J_bd").click()

                time.sleep(1)
                t = t - 1
            bd.minimize_window()
        except Exception as e:
            logging.error("Driver internal error, msg: %s." % e.msg)
            print("Driver internal error, msg: %s." % e.msg)  
        finally:
            g_sem.release()
            time.sleep(self.__sleep_sec)
            g_mutex.acquire()
            self.__bd_num = self.__bd_num - 1    
            g_mutex.release()  
            bd.quit()
        return

    def run(self):
        print("The primary thread is running...")
        while (True):
            print("Number of running work thread is %d" % self.__bd_num)
            if not(self.__bd_num < self.__cm.get_max_windou()):
                time.sleep(1)
                continue 
            #如果内置标志被置为False，主线程阻塞直到收到键盘运行事件
            g_run_event.wait()
            print("The primary thread has not recived the pause event...")
            #如果上个driver的启动未完成，主线程阻塞
            g_sem.acquire()
            print("The last open task has been finished...")

            th = Thread(target=self.once_search_task)
            #th.setDaemon(True)
            th.start()

class HotKey(Thread):
    def __init__(self,name):
        Thread.__init__(self)
        self.name = name

    def run(self):
        logging.info("Starting thread " + str(self.name))
        hot_key_main()
        logging.info("Stoping thread" + str(self.name))

 
def hot_key_main():
    user32 = ctypes.windll.user32
    while(True):
        #win+f9=run program win32con.MOD_WIN
        if not user32.RegisterHotKey(None, 98, 0, win32con.VK_F9):
            logging.error("Unable to register id 98 for run command.")
            print("Unable to register id 98 for run command.")
        #win+f10=pause program
        if not user32.RegisterHotKey(None, 99, 0, win32con.VK_F10):
            logging.error("Unable to register id 99 for pause command.")
            print("Unable to register id 99 for pause command.")
        try:
            msg = ctypes.wintypes.MSG()
            if user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    if msg.wParam == 98:
                        g_run_event.set()
                        logging.info("Program recived running event from keyboard...")
                        print("Program recived running event from keyboard...")
                    elif msg.wParam == 99:
                        g_run_event.clear()
                        logging.info("Program recived pause event from keyboard...")
                        print("Program recived pause event from keyboard...")

                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))
        finally:
            del msg
            user32.UnregisterHotKey(None, 98)
            user32.UnregisterHotKey(None, 99)

if __name__ == "__main__":
    try:
        c_service = Service("chromedriver.exe")
        c_service.command_line_args()
        c_service.start()
    except WebDriverException as e:
        logging.error("Failed to start broswer service")

    try:
        init_log_config()
        thread_hotKey = HotKey("thread_hotKey")
        thread_hotKey.setDaemon(True)
        thread_hotKey.start()
        print("The thread listening event from keyboard is running...")

        driver = Driver()
        driver.run()
    except KeyboardInterrupt as e:
        logging.warning("Handle KeyboardInterrupt.")
        c_service.stop()
        os.sys.exit(1)

