import ctypes
import json
import re
import subprocess
from os import environ
from pathlib import Path
from time import sleep
from typing import List

import psutil
import pyautogui
import win32clipboard
from pywinauto import Desktop
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.findwindows import find_elements as find_elements_low
from pywinauto.timings import wait_until_passes
from pywinauto.win32structures import RECT
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.switch_to import SwitchTo
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from win32gui import GetCursorInfo

D_TIMEOUT = 30
current_cursor = 0

if ctypes.windll.user32.GetKeyboardLayout(0) != 67699721:
    raise Exception('Смените раскладку на ENG')


class Json:
    @staticmethod
    def read(path):
        with open(str(path), 'r', encoding='utf-8') as fp:
            data = json.load(fp)
        return data

    @staticmethod
    def write(path, data):
        with open(str(path), 'w+', encoding='utf-8') as fp:
            json.dump(data, fp, ensure_ascii=False)


class Clipboard:
    @staticmethod
    def get():
        win32clipboard.OpenClipboard()
        result = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        if not len(result):
            raise Exception('Clipboard is empty')
        return result

    @staticmethod
    def set(value):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, value)
        win32clipboard.CloseClipboard()


class App:
    class Keys:
        CANCEL = '{VK_CANCEL}'  # ^break
        HELP = '{VK_HELP}'
        BACKSPACE = '{BACKSPACE}'
        BACK_SPACE = BACKSPACE
        TAB = '{VK_TAB}'
        CLEAR = '{VK_CLEAR}'
        RETURN = '{VK_RETURN}'
        ENTER = '{ENTER}'
        SHIFT = '{VK_LSHIFT}'
        LEFT_SHIFT = SHIFT
        CONTROL = '{VK_CONTROL}'
        LEFT_CONTROL = CONTROL
        ALT = '{VK_MENU}'
        LEFT_ALT = ALT
        PAUSE = '{VK_PAUSE}'
        ESCAPE = '{VK_ESCAPE}'
        SPACE = '{VK_SPACE}'
        PAGE_UP = '{PGUP}'
        PAGE_DOWN = '{PGDN}'
        END = '{VK_END}'
        HOME = '{VK_HOME}'
        LEFT = '{VK_LEFT}'
        ARROW_LEFT = LEFT
        UP = '{VK_UP}'
        ARROW_UP = UP
        RIGHT = '{VK_RIGHT}'
        ARROW_RIGHT = RIGHT
        DOWN = '{VK_DOWN}'
        ARROW_DOWN = DOWN
        INSERT = '{VK_INSERT}'
        DELETE = '{VK_DELETE}'

        NUMPAD0 = '{VK_NUMPAD0}'  # number pad keys
        NUMPAD1 = '{VK_NUMPAD1}'
        NUMPAD2 = '{VK_NUMPAD2}'
        NUMPAD3 = '{VK_NUMPAD3}'
        NUMPAD4 = '{VK_NUMPAD4}'
        NUMPAD5 = '{VK_NUMPAD5}'
        NUMPAD6 = '{VK_NUMPAD6}'
        NUMPAD7 = '{VK_NUMPAD7}'
        NUMPAD8 = '{VK_NUMPAD8}'
        NUMPAD9 = '{VK_NUMPAD9}'
        MULTIPLY = '{VK_MULTIPLY}'
        ADD = '{VK_ADD}'
        SEPARATOR = '{VK_SEPARATOR}'
        SUBTRACT = '{VK_SUBTRACT}'
        DECIMAL = '{VK_DECIMAL}'
        DIVIDE = '{VK_DIVIDE}'

        F1 = '{VK_F1}'  # function  keys
        F2 = '{VK_F2}'
        F3 = '{VK_F3}'
        F4 = '{VK_F4}'
        F5 = '{VK_F5}'
        F6 = '{VK_F6}'
        F7 = '{VK_F7}'
        F8 = '{VK_F8}'
        F9 = '{VK_F9}'
        F10 = '{VK_F10}'
        F11 = '{VK_F11}'
        F12 = '{VK_F12}'

        COMMAND = CONTROL

    keys = Keys

    @staticmethod
    def start_exe(program_path):
        return subprocess.Popen(program_path, stderr=subprocess.PIPE)

    @staticmethod
    def kill_exe(process_name, username=None, delay_before=0):
        username = username or f'{environ["userdomain"]}\\{environ.get("USERNAME")}'
        sleep(delay_before)
        process_list = [proc_ for proc_ in psutil.process_iter()]

        for process in process_list:
            try:
                if re.compile(process_name).match(process.name()) and username == process.username():
                    process.kill()
            except psutil.AccessDenied:
                continue

    @staticmethod
    def check_exe(process_name, username):
        flag_ = False
        process_list = [proc_ for proc_ in psutil.process_iter()]

        for process in process_list:
            try:
                if process.name() == process_name and process.username() == username:
                    flag_ = True
                    break
            except psutil.AccessDenied:
                continue
        return flag_

    @staticmethod
    def wait_cursor(timeout=float(D_TIMEOUT), appear=True):
        global current_cursor

        def main():
            global current_cursor
            current_cursor = GetCursorInfo()[1]
            flag = bool(current_cursor < 40000000 or current_cursor not in ['40306041', '56099749'])
            if flag != appear:
                raise Exception('Cursor code not appeared')

        try:
            wait_until_passes(timeout, 0.1, main)
            return True
        except (Exception,):
            return False

    @staticmethod
    def protect_str(value: str):
        except_list = ['(', ')', '+', '%', '^', '{', '}']
        replaced = []
        for char in [*value]:
            if char in except_list:
                char = '{' + char + '}'
            replaced.append(char)
        new_value = "".join(replaced)
        return new_value

    @staticmethod
    def find_elements_(timeout_, index=None, windows_=False, **kwargs):

        def get_windows_():
            # x, y = pyautogui.position()
            # pyautogui.moveTo(x + 1, y + 1)
            # pyautogui.moveTo(x, y)
            windows__ = Desktop(backend="uia").windows(**kwargs)
            if not len(windows__):
                raise Exception('Window not found')
            return windows__

        def get_elements_():
            # x, y = pyautogui.position()
            # pyautogui.moveTo(x + 1, y + 1)
            # pyautogui.moveTo(x, y)
            elements_ = find_elements_low(**kwargs, top_level_only=False)
            elements_ = [element_ for element_ in elements_ if element_.rectangle.left is not None]
            if not len(elements_):
                raise Exception('Element not found')
            return elements_

        elements__ = wait_until_passes(timeout_, 0.1, get_windows_ if windows_ else get_elements_)
        elements__: List[UIAWrapper] = elements__ if index is None else [elements__[index]]
        return elements__

    @classmethod
    def find_elements(cls, selector, timeout=D_TIMEOUT):
        cls.wait_cursor(D_TIMEOUT, appear=True)
        window_selector = dict(selector[0])
        window_rectangle = window_selector.get('rectangle')
        if window_rectangle:
            del window_selector['rectangle']
        windows = cls.find_elements_(timeout, windows_=True, **window_selector)
        if not len(windows):
            raise Exception('the matching window was not found')
        if window_rectangle:
            windows = [w for w in windows if w.element_info.rectangle == RECT(*window_rectangle)]
            if not len(windows):
                raise Exception('recheck the window rectangle')

        if len(selector) > 1:
            if len(windows) > 1:
                raise Exception(f'there are {len(windows)} windows, set window index in the selector')
            element_selector = dict(selector[1])
            element_rectangle = element_selector.get('rectangle')
            if element_rectangle:
                del element_selector['rectangle']
            elements = cls.find_elements_(timeout, **element_selector, parent=windows[0].element_info)
            elements = [UIAWrapper(e) for e in elements]
            if not len(elements):
                raise Exception('the matching element was not found')
            if element_rectangle:
                elements = [e for e in elements if e.element_info.rectangle == RECT(*element_rectangle)]
                if not len(elements):
                    raise Exception('recheck the element rectangle')
            return elements

        return windows

    @classmethod
    def find_element(cls, selector, timeout=D_TIMEOUT, index=0):
        return cls.find_elements(selector, timeout)[index]

    @classmethod
    def wait_element(cls, selector, timeout=D_TIMEOUT, appear=True):
        def main():
            # x, y = pyautogui.position()
            # pyautogui.moveTo(x + 1, y + 1)
            # pyautogui.moveTo(x, y)
            try:
                els = cls.find_elements(selector, 0)
                flag = bool(len(els))
            except (Exception,):
                flag = False

            if flag != appear:
                raise Exception('Element not appeared')

        try:
            wait_until_passes(timeout, 0.1, main)
            return True
        except (Exception,):
            return False


class Web:
    options = ChromeOptions
    keys = Keys
    by = By
    ec = expected_conditions
    ac = ActionChains

    def __init__(self, driver_path=None, download_path=None, user_data_dir=None, options=None, debug=False):
        default_driver_path = r"C:\Portable\PyCharmPortable\App\Chromium\chromedriver.exe"
        self.driver_path = driver_path if driver_path is not None else default_driver_path
        self.download_path = download_path if download_path is not None else Path.home().joinpath('Downloads').__str__()
        if options:
            self.options = options
        else:
            self.options = ChromeOptions()
            self.options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
            self.options.add_experimental_option("useAutomationExtension", False)
            self.options.add_experimental_option("prefs", {
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
                "profile.default_content_settings.popups": 0,
                "download.default_directory": self.download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": False,
                "profile.content_settings.exceptions.automatic_downloads.*.setting": 1
            })
            self.options.add_argument("--start-maximized")
            self.options.add_argument("--no-sandbox")
            self.options.add_argument("--disable-dev-shm-usage")
            self.options.add_argument("--disable-print-preview")
            self.options.add_argument("--disable-extensions")
            self.options.add_argument("--disable-notifications")
            self.options.add_argument("--ignore-ssl-errors=yes")
            self.options.add_argument("--ignore-certificate-errors")
        if user_data_dir:
            self.options.add_argument(f"user-data-dir={user_data_dir}")
        self.debug = debug

        self.webdriver = webdriver.Chrome(self.driver_path, options=self.options)
        self.switch_to = SwitchTo(self.webdriver)

    # ? Native

    def get(self, url: any):
        return self.webdriver.get(url)

    def back(self):
        return self.webdriver.back()

    def forward(self):
        return self.webdriver.forward()

    def close(self):
        return self.webdriver.close()

    def quit(self):
        return self.webdriver.quit()

    def maximize_window(self):
        return self.webdriver.maximize_window()

    def minimize_window(self):
        return self.webdriver.minimize_window()

    def execute_script(self, script: any, *args: any):
        return self.webdriver.execute_script(script, *args)

    def execute_async_script(self, script: any, *args: any):
        return self.webdriver.execute_async_script(script, *args)

    def save_screenshot(self, filename: str):
        return self.webdriver.save_screenshot(filename)

    def title(self):
        return self.webdriver.title

    def current_url(self):
        return self.webdriver.current_url

    def window_handles(self):
        return self.webdriver.window_handles

    def current_window_handle(self):
        return self.webdriver.current_window_handle

    # ? Extended

    def alert(self, text='', timeout=D_TIMEOUT):
        try:
            WebDriverWait(self.webdriver, timeout).until(expected_conditions.alert_is_present(), text)
            self.webdriver.switch_to.alert.accept()
            return True
        except TimeoutException:
            return False

    def focus(self, index=-1, frame=None, f_index=0):
        self.webdriver.switch_to.window(self.window_handles()[index])
        if frame:
            self.webdriver.switch_to.frame(self.find_elements(frame)[f_index].webobject)

    def find_element(self, selector, by='xpath', timeout=D_TIMEOUT, until_not=False, event=None):
        if event is None:
            event = expected_conditions.presence_of_element_located
        if timeout:
            self.wait_element(selector, by, timeout, until_not, event)
        webobject = self.webdriver.find_element(by, selector)
        el = WebElement(driver=self.webdriver, webobject=webobject, selector=selector, by=by, debug=self.debug)
        return el

    def find_elements(self, selector, by='xpath', timeout=D_TIMEOUT, until_not=False, event=None):
        if event is None:
            event = expected_conditions.presence_of_element_located
        if timeout:
            self.wait_element(selector, by, timeout, until_not, event)
        webobjects = self.webdriver.find_elements(by, selector)
        els = []
        for each in webobjects:
            els.append(WebElement(driver=self.webdriver, webobject=each, selector=selector, by=by, debug=self.debug))
        return els

    def wait_element(self, selector, by='xpath', timeout=D_TIMEOUT, until_not=False, event=None):
        if event is None:
            event = expected_conditions.presence_of_element_located
        try:
            if until_not:
                WebDriverWait(self.webdriver, timeout).until_not(event((by, selector)))
            else:
                WebDriverWait(self.webdriver, timeout).until(event((by, selector)))
            flag = True
        except Exception as e:
            if self.debug:
                print(e)
            flag = False
        if self.debug:
            print(selector, "appeared:", flag)
        return flag


class WebElement:
    class IsSelect:
        def __init__(self, webelement):
            self.__select = Select(webelement)
            self.options = self.__select.options
            self.all_selected_options = self.__select.all_selected_options
            self.is_multiple = self.__select.is_multiple

        # Related

        def select_by_index(self, index, delay_before=0, delay_after=0):
            sleep(delay_before)
            self.__select.select_by_index(index)
            sleep(delay_after)

        def select_by_value(self, value, delay_before=0, delay_after=0):
            sleep(delay_before)
            self.__select.select_by_value(value)
            sleep(delay_after)

        def select_by_visible_text(self, text, delay_before=0, delay_after=0):
            sleep(delay_before)
            self.__select.select_by_visible_text(text)
            sleep(delay_after)

        def deselect_by_index(self, index, delay_before=0, delay_after=0):
            sleep(delay_before)
            self.__select.deselect_by_index(index)
            sleep(delay_after)

        def deselect_by_value(self, value, delay_before=0, delay_after=0):
            sleep(delay_before)
            self.__select.deselect_by_value(value)
            sleep(delay_after)

        def deselect_by_visible_text(self, text, delay_before=0, delay_after=0):
            sleep(delay_before)
            self.__select.deselect_by_visible_text(text)
            sleep(delay_after)

        def deselect_all(self, delay_before=0, delay_after=0):
            sleep(delay_before)
            self.__select.deselect_all()
            sleep(delay_after)

    def __init__(self, driver, webobject, selector, by='xpath', debug=False):
        self.webdriver = driver
        self.webobject = webobject
        self.selector = selector
        self.by = by
        if webobject.tag_name.lower() == "select":
            self.is_select = self.IsSelect(self.webobject)
        self.debug = debug

    # ? Related

    def find_element(self, selector, by='xpath', timeout=D_TIMEOUT, until_not=False, event=None):
        if event is None:
            event = expected_conditions.presence_of_element_located
        if selector[0] != '.':
            raise ValueError("Дочерний селектор должен начинаться с '.'")
        if timeout:
            self.wait_element(selector, by, timeout, until_not, event)
        webobject = self.webobject.find_element(by, selector)
        el = WebElement(driver=self.webdriver, webobject=webobject, selector=selector, by=by, debug=self.debug)
        return el

    def find_elements(self, selector, by='xpath', timeout=D_TIMEOUT, until_not=False, event=None):
        if event is None:
            event = expected_conditions.presence_of_element_located
        if selector[0] != '.':
            raise ValueError("Дочерний селектор должен начинаться с '.'")
        if timeout:
            self.wait_element(selector, by, timeout, until_not, event)
        webobjects = self.webobject.find_elements(by, selector)
        els = []
        for each in webobjects:
            els.append(WebElement(driver=self.webdriver, webobject=each, selector=selector, by=by, debug=self.debug))
        return els

    def wait_element(self, selector, by='xpath', timeout=D_TIMEOUT, until_not=False, event=None, exc=True):
        if selector[0] == '.':
            selector = f'{self.selector}{selector[1:]}'
        else:
            selector = f'{self.selector}{selector}'
            if exc:
                raise ValueError("Дочерний селектор должен начинаться с '.'")

        if event is None:
            event = expected_conditions.presence_of_element_located
        try:
            if until_not:
                WebDriverWait(self.webdriver, timeout).until_not(event((by, selector)))
            else:
                WebDriverWait(self.webdriver, timeout).until(event((by, selector)))
            flag = True
        except Exception as e:
            if self.debug:
                print(e)
            flag = False
        if self.debug:
            print(selector, "appeared:", flag)
        return flag

    # ? Actions

    def click(self, delay_before=0, delay_after=0, scroll=True):
        sleep(delay_before)
        if scroll:
            self.scroll()
        self.webobject.click()
        sleep(delay_after)
        return self

    def double_click(self, delay_before=0, delay_after=0, scroll=True):
        sleep(delay_before)
        if scroll:
            self.scroll()
        ac = ActionChains(self.webdriver)
        ac.double_click(self.webobject).perform()
        sleep(delay_after)
        return self

    def scroll(self, delay_before=0, delay_after=0):
        sleep(delay_before)
        try:
            ac = ActionChains(self.webdriver)
            ac.move_to_element(self.webobject).perform()
        except (Exception,):
            pass
        sleep(delay_after)
        return self

    def send_keys(self, *args, delay_before=0, delay_after=0, clear=False):
        sleep(delay_before)
        if clear:
            self.webobject.clear()
        self.webobject.send_keys(*args)
        sleep(delay_after)
        return self

    def get_attribute(self, name, delay_before=0, delay_after=0):
        sleep(delay_before)
        value = self.webobject.get_attribute(name)
        sleep(delay_after)
        return value

    def get_text(self, delay_before=0, delay_after=0):
        sleep(delay_before)
        value = self.webobject.text
        sleep(delay_after)
        return str(value)

    def clear(self, delay_before=0, delay_after=0):
        sleep(delay_before)
        self.webobject.clear()
        sleep(delay_after)
        return self


if __name__ == '__main__':
    pass
